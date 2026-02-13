/*
 * drill_render.c -- GDI owner-drawn panel for pronunciation drill
 *
 * Custom child window: target sentence with inline result below each char,
 * timing visualization, hesitation highlighting, per-char duration labels.
 * Click target row to copy original, click result row to copy transcription.
 * #included from main.c (voice-test-gui), after drill.c.
 *
 * Layout:
 *   [HSK N]
 *   Target chars (large, individual cells) -- click to copy original
 *   Result chars (same size, state-dependent) -- click to copy transcription
 *   Timing bars (thin colored bars showing inter-character gaps)
 *   Duration labels (ms per character, small font)
 *   Pinyin
 *   English
 *   Status line (feedback + stats + speaking pace)
 */

#define DRILL_COLOR_BG         RGB(30, 30, 30)
#define DRILL_COLOR_TEXT       RGB(240, 240, 240)
#define DRILL_COLOR_PINYIN     RGB(160, 160, 170)
#define DRILL_COLOR_ENGLISH    RGB(130, 130, 140)
#define DRILL_COLOR_MATCH_BG   RGB(40, 120, 40)
#define DRILL_COLOR_MISS_BG    RGB(140, 40, 40)
#define DRILL_COLOR_MATCH_FG   RGB(220, 255, 220)
#define DRILL_COLOR_MISS_FG    RGB(255, 200, 200)
#define DRILL_COLOR_STATUS     RGB(180, 180, 190)
#define DRILL_COLOR_CORRECT    RGB(80, 200, 80)
#define DRILL_COLOR_HSK_LABEL  RGB(100, 140, 200)

/* Indicator colors for result cell states */
#define DRILL_COLOR_IDLE       RGB(50, 50, 55)
#define DRILL_COLOR_RECORDING  RGB(180, 140, 30)
#define DRILL_COLOR_PENDING    RGB(50, 90, 160)
#define DRILL_COLOR_STREAM_FG  RGB(200, 200, 210)

/* Timing bar colors */
#define DRILL_COLOR_TIME_FAST  RGB(60, 160, 60)
#define DRILL_COLOR_TIME_MED   RGB(140, 140, 50)
#define DRILL_COLOR_TIME_SLOW  RGB(180, 80, 30)
#define DRILL_COLOR_TIME_DUR   RGB(120, 120, 130)

/* Hesitation threshold (ms between consecutive chars) */
#define DRILL_HESITATE_MS      500
#define DRILL_COLOR_HESITATE   RGB(200, 160, 40)

/* Timing bar geometry */
#define DRILL_TIMEBAR_H        6
#define DRILL_TIMEBAR_PAD      2

/* Copy flash overlay */
#define DRILL_COLOR_COPY_BG    RGB(60, 60, 70)
#define DRILL_COLOR_COPY_FG    RGB(200, 220, 200)

/* Row Y positions for click detection (set during WM_PAINT) */
static int s_target_y_top = 0, s_target_y_bot = 0;
static int s_result_y_top = 0, s_result_y_bot = 0;

/* Copy UTF-8 string to clipboard as Unicode */
static void drill_copy_to_clipboard(HWND hwnd, const char *utf8)
{
    if (!utf8 || !utf8[0]) return;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) return;
    HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, (SIZE_T)wlen * sizeof(wchar_t));
    if (!hg) return;
    wchar_t *dst = (wchar_t *)GlobalLock(hg);
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, dst, wlen);
    GlobalUnlock(hg);
    if (OpenClipboard(hwnd)) {
        EmptyClipboard();
        SetClipboardData(CF_UNICODETEXT, hg);
        CloseClipboard();
    } else {
        GlobalFree(hg);
    }
}

/* Build UTF-8 string from codepoint array */
static int drill_cps_to_utf8(const int *cps, int n, char *buf, int buf_size)
{
    int pos = 0;
    for (int i = 0; i < n && pos < buf_size - 4; i++) {
        pos += cp_to_utf8(cps[i], buf + pos);
    }
    buf[pos] = '\0';
    return pos;
}

/* Draw a single codepoint in a cell */
static void drill_draw_cp(HDC hdc, int cp, RECT *rc, COLORREF fg)
{
    if (cp <= 0) return;
    wchar_t wc[3] = {0};
    if (cp <= 0xFFFF) {
        wc[0] = (wchar_t)cp;
    } else {
        cp -= 0x10000;
        wc[0] = (wchar_t)(0xD800 + (cp >> 10));
        wc[1] = (wchar_t)(0xDC00 + (cp & 0x3FF));
    }
    SetTextColor(hdc, fg);
    DrawTextW(hdc, wc, -1, rc, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
}

/* Fill a small indicator rectangle centered in a cell */
static void drill_draw_indicator(HDC hdc, RECT *cell, COLORREF color)
{
    int cx = (cell->left + cell->right) / 2;
    int cy = (cell->top + cell->bottom) / 2;
    int sz = 6;
    RECT dot = { cx - sz, cy - sz, cx + sz, cy + sz };
    HBRUSH br = CreateSolidBrush(color);
    FillRect(hdc, &dot, br);
    DeleteObject(br);
}

/* Color for a timing delta (ms between consecutive chars) */
static COLORREF drill_time_color(int delta_ms)
{
    if (delta_ms < 200) return DRILL_COLOR_TIME_FAST;
    if (delta_ms < DRILL_HESITATE_MS) return DRILL_COLOR_TIME_MED;
    return DRILL_COLOR_TIME_SLOW;
}

/* Draw timing bars and duration labels for a row of characters.
 * ms[] has timing for each position, n = count.
 * Returns vertical space consumed (y advance). */
static int drill_render_timing(HDC hdc, int *ms, int n,
                               int start_x, int cell_w, int y, int w)
{
    if (n < 2) return 0;
    (void)w;

    int y_bar = y + DRILL_TIMEBAR_PAD;

    /* Compute deltas and find max for scaling */
    int max_delta = 1;
    for (int i = 1; i < n; i++) {
        int d = ms[i] - ms[i - 1];
        if (d < 0) d = 0;
        if (d > max_delta) max_delta = d;
    }

    /* Draw timing bars: width proportional to delta, colored by speed */
    for (int i = 0; i < n; i++) {
        int delta = (i == 0) ? 0 : (ms[i] - ms[i - 1]);
        if (delta < 0) delta = 0;

        /* Bar width: proportional to delta, clamped to cell width */
        int bar_w = (max_delta > 0) ? (delta * (cell_w - 4)) / max_delta : 0;
        if (bar_w < 2 && delta > 0) bar_w = 2;

        int cx = start_x + i * cell_w + cell_w / 2;
        RECT bar_rc = { cx - bar_w / 2, y_bar, cx + bar_w / 2, y_bar + DRILL_TIMEBAR_H };

        COLORREF col = drill_time_color(delta);
        HBRUSH br = CreateSolidBrush(col);
        FillRect(hdc, &bar_rc, br);
        DeleteObject(br);

        /* Hesitation border on the bar area */
        if (delta >= DRILL_HESITATE_MS) {
            HPEN pen = CreatePen(PS_SOLID, 2, DRILL_COLOR_HESITATE);
            HPEN old_pen = (HPEN)SelectObject(hdc, pen);
            RECT highlight = { start_x + i * cell_w + 1, y_bar - 1,
                               start_x + (i + 1) * cell_w - 1, y_bar + DRILL_TIMEBAR_H + 1 };
            HBRUSH null_br = (HBRUSH)GetStockObject(NULL_BRUSH);
            SelectObject(hdc, null_br);
            Rectangle(hdc, highlight.left, highlight.top, highlight.right, highlight.bottom);
            SelectObject(hdc, old_pen);
            DeleteObject(pen);
        }
    }

    int y_labels = y_bar + DRILL_TIMEBAR_H + 2;

    /* Duration labels (ms) under each cell */
    SelectObject(hdc, g_font_small);
    SetTextColor(hdc, DRILL_COLOR_TIME_DUR);
    for (int i = 0; i < n; i++) {
        int delta = (i == 0) ? ms[0] : (ms[i] - ms[i - 1]);
        if (delta < 0) delta = 0;
        char label[16];
        snprintf(label, sizeof(label), "%dms", delta);
        RECT lrc = { start_x + i * cell_w, y_labels,
                     start_x + (i + 1) * cell_w, y_labels + 14 };
        DrawTextA(hdc, label, -1, &lrc, DT_CENTER | DT_SINGLELINE);
    }

    return DRILL_TIMEBAR_PAD + DRILL_TIMEBAR_H + 2 + 14 + 2;
}

/* Draw "Copied!" overlay on a row */
static void drill_draw_copy_overlay(HDC hdc, int y_top, int y_bot, int w, int margin)
{
    RECT overlay = { margin, y_top, w - margin, y_bot };
    HBRUSH br = CreateSolidBrush(DRILL_COLOR_COPY_BG);
    FillRect(hdc, &overlay, br);
    DeleteObject(br);
    SelectObject(hdc, g_font_medium);
    SetTextColor(hdc, DRILL_COLOR_COPY_FG);
    DrawTextA(hdc, "Copied!", -1, &overlay, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
}

static LRESULT CALLBACK DrillWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc_screen = BeginPaint(hwnd, &ps);

        RECT rc;
        GetClientRect(hwnd, &rc);
        int w = rc.right - rc.left;
        int h = rc.bottom - rc.top;

        /* Double buffering */
        HDC hdc = CreateCompatibleDC(hdc_screen);
        HBITMAP bmp = CreateCompatibleBitmap(hdc_screen, w, h);
        HBITMAP old_bmp = (HBITMAP)SelectObject(hdc, bmp);

        /* Background */
        HBRUSH bg_brush = CreateSolidBrush(DRILL_COLOR_BG);
        FillRect(hdc, &rc, bg_brush);
        DeleteObject(bg_brush);
        SetBkMode(hdc, TRANSPARENT);

        if (g_drill_state.current_idx < 0 || g_drill_state.num_sentences == 0) {
            SelectObject(hdc, g_font_medium);
            SetTextColor(hdc, DRILL_COLOR_STATUS);
            RECT text_rc = { 20, h / 2 - 20, w - 20, h / 2 + 20 };
            DrawTextA(hdc, "No sentences loaded", -1, &text_rc,
                      DT_CENTER | DT_SINGLELINE | DT_VCENTER);
            goto paint_done;
        }

        DrillSentence *sent = &g_drill_state.sentences[g_drill_state.current_idx];
        DrillSentenceProgress *prog = &g_drill_state.progress[g_drill_state.current_idx];
        DrillDiff *diff = &g_drill_state.last_diff;
        int y = 12;
        int margin = 20;

        /* HSK level badge */
        {
            char hsk_label[32];
            snprintf(hsk_label, sizeof(hsk_label), "HSK %d", sent->hsk_level);
            SelectObject(hdc, g_font_medium);
            SetTextColor(hdc, DRILL_COLOR_HSK_LABEL);
            RECT hsk_rc = { margin, y, w - margin, y + 22 };
            DrawTextA(hdc, hsk_label, -1, &hsk_rc, DT_CENTER | DT_SINGLELINE);
            y += 24;
        }

        /* Extract target codepoints (strip punctuation) */
        int target_cps[DRILL_MAX_TEXT];
        int num_target = utf8_to_codepoints(sent->chinese, target_cps, DRILL_MAX_TEXT);
        num_target = strip_codepoints(target_cps, num_target);

        /* Measure cell size using the large Chinese font */
        SelectObject(hdc, g_font_drill_chinese);
        SIZE char_size;
        GetTextExtentPoint32W(hdc, L"\x4F60", 1, &char_size);
        int cell_w = char_size.cx + 6;
        int cell_h = char_size.cy + 4;

        /* Determine how many cells (target + possible excess from result) */
        int num_result_extra = 0;
        if (g_drill_state.has_result && diff->num_actual > num_target)
            num_result_extra = diff->num_actual - num_target;
        else if (!g_drill_state.has_result && g_drill_stream_len > num_target)
            num_result_extra = g_drill_stream_len - num_target;
        int total_cols = num_target + num_result_extra;

        int total_w = total_cols * cell_w;
        int start_x = (w - total_w) / 2;
        if (start_x < margin) start_x = margin;

        /* Row 1: Target characters (large font, white) */
        s_target_y_top = y;
        SelectObject(hdc, g_font_drill_chinese);
        for (int i = 0; i < num_target; i++) {
            RECT cell = { start_x + i * cell_w, y,
                          start_x + (i + 1) * cell_w, y + cell_h };
            drill_draw_cp(hdc, target_cps[i], &cell, DRILL_COLOR_TEXT);
        }
        y += cell_h + 2;
        s_target_y_bot = y;

        /* Copy flash overlay for target row */
        if (g_drill_copy_row == 0) {
            drill_draw_copy_overlay(hdc, s_target_y_top, s_target_y_bot, w, margin);
        }

        /* Row 2: Result characters (same size, directly below target) */
        s_result_y_top = y;

        /* Determine drill phase */
        int has_result = g_drill_state.has_result;
        int phase;
        if (has_result) {
            phase = 4;
        } else if (g_drill_stream_len > 0) {
            phase = 3;
        } else if (g_is_recording) {
            phase = 1;
        } else if (g_transcribe_thread != NULL) {
            phase = 2;
        } else {
            phase = 0;
        }

        /* In final phase, the number of meaningful result cells is
         * max(actual, expected). Don't draw empty cells beyond that. */
        int result_cols = total_cols;
        if (phase == 4) {
            int max_diff = diff->num_actual > diff->num_expected
                         ? diff->num_actual : diff->num_expected;
            if (max_diff < result_cols) result_cols = max_diff;
        }

        /* Track whether we have timing data to show */
        int has_timing = 0;
        int timing_n = 0;

        {
            SelectObject(hdc, g_font_drill_chinese);

            for (int i = 0; i < result_cols; i++) {
                RECT cell = { start_x + i * cell_w, y,
                              start_x + (i + 1) * cell_w, y + cell_h };

                if (phase == 4) {
                    /* Final result: colored background + character */
                    COLORREF bg_col;
                    if (i < diff->num_actual && i < diff->num_expected) {
                        bg_col = diff->char_match[i] ? DRILL_COLOR_MATCH_BG : DRILL_COLOR_MISS_BG;
                    } else {
                        bg_col = DRILL_COLOR_MISS_BG;
                    }

                    /* Hesitation highlight: amber border on cells with large gaps */
                    if (i > 0 && i < g_drill_stream_len && g_drill_stream_len >= 2) {
                        int delta = g_drill_stream_ms[i] - g_drill_stream_ms[i - 1];
                        if (delta >= DRILL_HESITATE_MS) {
                            RECT border = { cell.left - 1, cell.top - 1,
                                            cell.right + 1, cell.bottom + 1 };
                            HBRUSH hbr = CreateSolidBrush(DRILL_COLOR_HESITATE);
                            FillRect(hdc, &border, hbr);
                            DeleteObject(hbr);
                        }
                    }

                    HBRUSH br = CreateSolidBrush(bg_col);
                    FillRect(hdc, &cell, br);
                    DeleteObject(br);

                    int cp = -1;
                    COLORREF fg;
                    if (i < diff->num_actual) {
                        cp = diff->actual_cps[i];
                        fg = (i < diff->num_expected && diff->char_match[i])
                            ? DRILL_COLOR_MATCH_FG : DRILL_COLOR_MISS_FG;
                    } else if (i < diff->num_expected) {
                        cp = diff->expected_cps[i];
                        fg = RGB(100, 60, 60);
                    }
                    if (cp > 0) drill_draw_cp(hdc, cp, &cell, fg);

                } else if (phase == 3 && i < g_drill_stream_len) {
                    /* Streaming: show character as it arrives */
                    drill_draw_cp(hdc, g_drill_stream_cps[i], &cell, DRILL_COLOR_STREAM_FG);

                } else if (phase == 3 && i >= g_drill_stream_len && i < num_target) {
                    drill_draw_indicator(hdc, &cell, DRILL_COLOR_PENDING);

                } else if (phase == 2) {
                    drill_draw_indicator(hdc, &cell, DRILL_COLOR_PENDING);

                } else if (phase == 1) {
                    drill_draw_indicator(hdc, &cell, DRILL_COLOR_RECORDING);

                } else {
                    drill_draw_indicator(hdc, &cell, DRILL_COLOR_IDLE);
                }
            }
            y += cell_h + 2;
            s_result_y_bot = y;

            /* Timing: clamp to visible result columns */
            if (g_drill_stream_len >= 2) {
                has_timing = 1;
                timing_n = g_drill_stream_len;
                if (timing_n > result_cols) timing_n = result_cols;
            }
        }

        /* Copy flash overlay for result row */
        if (g_drill_copy_row == 1) {
            drill_draw_copy_overlay(hdc, s_result_y_top, s_result_y_bot, w, margin);
        }

        /* Timing visualization (bars + duration labels) */
        if (has_timing && timing_n >= 2 && (phase == 3 || phase == 4)) {
            y += drill_render_timing(hdc, g_drill_stream_ms, timing_n,
                                     start_x, cell_w, y, w);
        }

        /* Pinyin */
        {
            SelectObject(hdc, g_font_medium);
            SetTextColor(hdc, DRILL_COLOR_PINYIN);
            RECT pin_rc = { margin, y, w - margin, y + 24 };
            DrawTextA(hdc, sent->pinyin, -1, &pin_rc,
                      DT_CENTER | DT_SINGLELINE | DT_VCENTER);
            y += 26;
        }

        /* English */
        {
            SelectObject(hdc, g_font_medium);
            SetTextColor(hdc, DRILL_COLOR_ENGLISH);
            RECT eng_rc = { margin, y, w - margin, y + 24 };
            DrawTextA(hdc, sent->english, -1, &eng_rc,
                      DT_CENTER | DT_SINGLELINE | DT_VCENTER);
        }

        /* Status line at bottom */
        {
            int status_y = rc.bottom - 36;
            SelectObject(hdc, g_font_medium);

            /* Left: result feedback */
            if (g_drill_state.has_result) {
                const char *feedback;
                COLORREF feedback_color;
                if (g_drill_state.last_diff.match) {
                    feedback = "Correct! Space for next";
                    feedback_color = DRILL_COLOR_CORRECT;
                } else {
                    feedback = "Try again";
                    feedback_color = RGB(255, 100, 100);
                }
                SetTextColor(hdc, feedback_color);
                RECT fb_rc = { margin, status_y, w / 3, status_y + 24 };
                DrawTextA(hdc, feedback, -1, &fb_rc, DT_LEFT | DT_SINGLELINE | DT_VCENTER);
            }

            /* Center: speaking pace (chars/sec) */
            if (has_timing && timing_n >= 2) {
                int total_ms = g_drill_stream_ms[timing_n - 1] - g_drill_stream_ms[0];
                if (total_ms > 0) {
                    double cps = (double)(timing_n - 1) * 1000.0 / (double)total_ms;
                    char pace[48];
                    snprintf(pace, sizeof(pace), "%.1f char/s", cps);
                    SetTextColor(hdc, DRILL_COLOR_PINYIN);
                    RECT pace_rc = { w / 3, status_y, 2 * w / 3, status_y + 24 };
                    DrawTextA(hdc, pace, -1, &pace_rc, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
                }
            }

            /* Right: session stats */
            char stats[128];
            if (g_drill_state.session_attempts > 0) {
                int pct = (g_drill_state.session_correct * 100) / g_drill_state.session_attempts;
                snprintf(stats, sizeof(stats), "%d/%d (%d%%)",
                         g_drill_state.session_correct, g_drill_state.session_attempts, pct);
            } else {
                snprintf(stats, sizeof(stats), "D:drill  H:HSK filter");
            }
            SetTextColor(hdc, DRILL_COLOR_STATUS);
            RECT st_rc = { 2 * w / 3, status_y, w - margin, status_y + 24 };
            DrawTextA(hdc, stats, -1, &st_rc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

            /* Sentence progress (if has attempts) */
            if (prog->attempts > 0) {
                char sprog[64];
                snprintf(sprog, sizeof(sprog), "This: %d/%d  Streak: %d",
                         prog->correct, prog->attempts, prog->streak);
                SetTextColor(hdc, DRILL_COLOR_PINYIN);
                RECT sp_rc = { margin, status_y - 24, w - margin, status_y };
                DrawTextA(hdc, sprog, -1, &sp_rc, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
            }
        }

paint_done:
        /* Blit to screen */
        BitBlt(hdc_screen, 0, 0, w, h, hdc, 0, 0, SRCCOPY);
        SelectObject(hdc, old_bmp);
        DeleteObject(bmp);
        DeleteDC(hdc);

        EndPaint(hwnd, &ps);
        return 0;
    }

    case WM_ERASEBKGND:
        return 1;

    case WM_LBUTTONUP: {
        int click_y = HIWORD(lParam);
        if (g_drill_state.current_idx < 0 ||
            g_drill_state.current_idx >= g_drill_state.num_sentences)
            return 0;

        DrillSentence *s = &g_drill_state.sentences[g_drill_state.current_idx];

        if (click_y >= s_target_y_top && click_y < s_target_y_bot) {
            /* Click on target row: copy original Chinese */
            drill_copy_to_clipboard(hwnd, s->chinese);
            log_event("DRILL", "Copied target Chinese to clipboard");
            g_drill_copy_row = 0;
            g_drill_copy_tick = GetTickCount();
            SetTimer(g_hwnd_main, ID_TIMER_DRILL_COPY, DRILL_COPY_FLASH_MS, NULL);
            InvalidateRect(hwnd, NULL, FALSE);
        }
        else if (click_y >= s_result_y_top && click_y < s_result_y_bot) {
            /* Click on result row: copy transcription */
            const char *text = NULL;
            char stream_buf[1024];
            if (g_drill_state.has_result && g_drill_state.result_text[0]) {
                text = g_drill_state.result_text;
            } else if (g_drill_stream_len > 0) {
                drill_cps_to_utf8(g_drill_stream_cps, g_drill_stream_len,
                                  stream_buf, sizeof(stream_buf));
                text = stream_buf;
            }
            if (text) {
                drill_copy_to_clipboard(hwnd, text);
                log_event("DRILL", "Copied transcription to clipboard");
                g_drill_copy_row = 1;
                g_drill_copy_tick = GetTickCount();
                SetTimer(g_hwnd_main, ID_TIMER_DRILL_COPY, DRILL_COPY_FLASH_MS, NULL);
                InvalidateRect(hwnd, NULL, FALSE);
            }
        }
        return 0;
    }

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
}

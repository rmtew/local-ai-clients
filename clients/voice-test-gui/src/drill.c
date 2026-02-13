/*
 * drill.c -- Pronunciation drill core logic
 *
 * Sentence loading, UTF-8 diff, weighted selection, progress I/O.
 * #included from main.c (voice-test-gui).
 */

/* ---------- UTF-8 helpers ---------- */

/* Decode one UTF-8 codepoint from str. Returns bytes consumed (1-4), or 0 on error. */
static int utf8_next_cp(const char *str, int *cp)
{
    const unsigned char *s = (const unsigned char *)str;
    if (s[0] == 0) { *cp = 0; return 0; }

    if (s[0] < 0x80) {
        *cp = s[0];
        return 1;
    }
    if ((s[0] & 0xE0) == 0xC0) {
        if ((s[1] & 0xC0) != 0x80) { *cp = 0xFFFD; return 1; }
        *cp = ((s[0] & 0x1F) << 6) | (s[1] & 0x3F);
        return 2;
    }
    if ((s[0] & 0xF0) == 0xE0) {
        if ((s[1] & 0xC0) != 0x80 || (s[2] & 0xC0) != 0x80) { *cp = 0xFFFD; return 1; }
        *cp = ((s[0] & 0x0F) << 12) | ((s[1] & 0x3F) << 6) | (s[2] & 0x3F);
        return 3;
    }
    if ((s[0] & 0xF8) == 0xF0) {
        if ((s[1] & 0xC0) != 0x80 || (s[2] & 0xC0) != 0x80 || (s[3] & 0xC0) != 0x80) {
            *cp = 0xFFFD; return 1;
        }
        *cp = ((s[0] & 0x07) << 18) | ((s[1] & 0x3F) << 12) |
              ((s[2] & 0x3F) << 6) | (s[3] & 0x3F);
        return 4;
    }
    *cp = 0xFFFD;
    return 1;
}

/* Encode one codepoint to UTF-8. Returns bytes written (1-4). */
static int cp_to_utf8(int cp, char *buf)
{
    unsigned char *b = (unsigned char *)buf;
    if (cp < 0x80) {
        b[0] = (unsigned char)cp;
        return 1;
    }
    if (cp < 0x800) {
        b[0] = 0xC0 | (cp >> 6);
        b[1] = 0x80 | (cp & 0x3F);
        return 2;
    }
    if (cp < 0x10000) {
        b[0] = 0xE0 | (cp >> 12);
        b[1] = 0x80 | ((cp >> 6) & 0x3F);
        b[2] = 0x80 | (cp & 0x3F);
        return 3;
    }
    b[0] = 0xF0 | (cp >> 18);
    b[1] = 0x80 | ((cp >> 12) & 0x3F);
    b[2] = 0x80 | ((cp >> 6) & 0x3F);
    b[3] = 0x80 | (cp & 0x3F);
    return 4;
}

/* Extract codepoints from UTF-8 string into array. Returns count. */
static int utf8_to_codepoints(const char *str, int *cps, int max_cps)
{
    int count = 0;
    const char *p = str;
    while (*p && count < max_cps) {
        int cp;
        int bytes = utf8_next_cp(p, &cp);
        if (bytes == 0) break;
        cps[count++] = cp;
        p += bytes;
    }
    return count;
}

/* Check if codepoint is Chinese punctuation or ASCII whitespace/punctuation to strip */
static int is_strip_cp(int cp)
{
    if (cp <= 0x20) return 1;                           /* ASCII control + space */
    if (cp == '.' || cp == ',' || cp == '!' || cp == '?' || cp == ';') return 1;
    /* Chinese punctuation */
    if (cp == 0x3002) return 1;  /* Ideographic full stop */
    if (cp == 0xFF0C) return 1;  /* Fullwidth comma */
    if (cp == 0xFF01) return 1;  /* Fullwidth exclamation */
    if (cp == 0xFF1F) return 1;  /* Fullwidth question */
    if (cp == 0x3001) return 1;  /* Ideographic comma */
    if (cp == 0xFF1B) return 1;  /* Fullwidth semicolon */
    if (cp == 0x2026) return 1;  /* Ellipsis */
    if (cp == 0x300A || cp == 0x300B) return 1;  /* Angle brackets */
    if (cp == 0x201C || cp == 0x201D) return 1;  /* Smart double quotes */
    if (cp == 0x2018 || cp == 0x2019) return 1;  /* Smart single quotes */
    return 0;
}

/* Strip leading/trailing whitespace and punctuation from codepoint array in-place.
   Returns new count. */
static int strip_codepoints(int *cps, int count)
{
    int start = 0, end = count;
    while (start < end && is_strip_cp(cps[start])) start++;
    while (end > start && is_strip_cp(cps[end - 1])) end--;
    if (start > 0) {
        for (int i = 0; i < end - start; i++)
            cps[i] = cps[start + i];
    }
    return end - start;
}

/* ---------- Homophone equivalence ---------- */

/* Groups of characters with identical pronunciation.  Characters within
 * the same group are treated as matching in drill_check(). */
static const int HOMOPHONES[][6] = {
    /* ta1 */  {0x4ED6, 0x5979, 0x5B83, 0},           /* he she it */
    /* de */   {0x7684, 0x5730, 0x5F97, 0},           /* de/di de de */
    /* ta1men */ {0x4EEC, 0},                          /* (just in case) */
    /* zhe4 */ {0x8FD9, 0x9019, 0},                   /* simplified/traditional */
    /* na4 */  {0x90A3, 0x5462, 0},                   /* that / particle */
    /* ma */   {0x5417, 0x5440, 0x561B, 0x55CE, 0},   /* question particles */
    /* le */   {0x4E86, 0x4E86, 0},
    /* shi4 */ {0x662F, 0x4E8B, 0},                   /* is / matter */
    /* zai4 */ {0x5728, 0x518D, 0},                   /* at / again */
    /* ji3 */  {0x51E0, 0x5E7E, 0},                   /* simplified/traditional */
    /* dian3 */ {0x70B9, 0x9EDE, 0},                  /* simplified/traditional */
    /* li3 */  {0x91CC, 0x88E1, 0x88CF, 0},           /* inside variants */
    /* hao3 */ {0x597D, 0},
    /* xiang3/xiang1 */ {0x60F3, 0x76F8, 0},
    /* guo2/guo4 */ {0x56FD, 0x570B, 0x8FC7, 0x904E, 0}, /* country/pass simp/trad */
    /* dou1 */ {0x90FD, 0},
    /* wei4/wei2 */ {0x4E3A, 0x70BA, 0},              /* simplified/traditional */
    /* shen2me */ {0x4EC0, 0x751A, 0},
    /* me/mo */ {0x4E48, 0x9EBC, 0x9EBD, 0},          /* simplified/traditional */
    /* hui4 */ {0x4F1A, 0x6703, 0},                   /* simplified/traditional */
    /* lv3 */  {0x65C5, 0},
    /* you2 */ {0x6E38, 0x904A, 0},                   /* simplified/traditional */
    {0}  /* sentinel */
};

/* Check if two codepoints are homophone-equivalent */
static int homophones_match(int cp_a, int cp_b)
{
    if (cp_a == cp_b) return 1;
    for (int g = 0; HOMOPHONES[g][0] != 0; g++) {
        int found_a = 0, found_b = 0;
        for (int i = 0; HOMOPHONES[g][i] != 0 && i < 6; i++) {
            if (HOMOPHONES[g][i] == cp_a) found_a = 1;
            if (HOMOPHONES[g][i] == cp_b) found_b = 1;
        }
        if (found_a && found_b) return 1;
    }
    return 0;
}

/* ---------- Sentence bank loading ---------- */

static int drill_load_sentences(DrillState *ds, const char *filepath)
{
    FILE *f = fopen(filepath, "r");
    if (!f) return -1;

    char line[1024];
    int current_hsk = 1;
    ds->num_sentences = 0;

    while (fgets(line, sizeof(line), f) && ds->num_sentences < DRILL_MAX_SENTENCES) {
        /* Strip trailing newline */
        int len = (int)strlen(line);
        while (len > 0 && (line[len - 1] == '\n' || line[len - 1] == '\r'))
            line[--len] = '\0';

        /* Skip empty lines */
        if (len == 0) continue;

        /* HSK level header */
        if (line[0] == '#') {
            const char *p = line + 1;
            while (*p == ' ') p++;
            if (_strnicmp(p, "HSK", 3) == 0) {
                p += 3;
                while (*p == ' ') p++;
                current_hsk = atoi(p);
                if (current_hsk < 1) current_hsk = 1;
                if (current_hsk > 6) current_hsk = 6;
            }
            continue;
        }

        /* Parse pipe-delimited: chinese|pinyin|english */
        char *p1 = strchr(line, '|');
        if (!p1) continue;
        *p1 = '\0';
        char *p2 = strchr(p1 + 1, '|');
        if (!p2) continue;
        *p2 = '\0';

        DrillSentence *s = &ds->sentences[ds->num_sentences];
        strncpy(s->chinese, line, DRILL_MAX_TEXT - 1);
        s->chinese[DRILL_MAX_TEXT - 1] = '\0';
        strncpy(s->pinyin, p1 + 1, DRILL_MAX_TEXT - 1);
        s->pinyin[DRILL_MAX_TEXT - 1] = '\0';
        strncpy(s->english, p2 + 1, DRILL_MAX_TEXT - 1);
        s->english[DRILL_MAX_TEXT - 1] = '\0';
        s->hsk_level = current_hsk;

        ds->num_sentences++;
    }

    fclose(f);
    return (ds->num_sentences > 0) ? 0 : -1;
}

/* ---------- Progress I/O ---------- */

/* Load progress from tab-separated file. Format: chinese_chars\tattempts\tcorrect\tstreak */
static void drill_load_progress(DrillState *ds, const char *filepath)
{
    FILE *f = fopen(filepath, "r");
    if (!f) return;

    char line[1024];
    while (fgets(line, sizeof(line), f)) {
        int len = (int)strlen(line);
        while (len > 0 && (line[len - 1] == '\n' || line[len - 1] == '\r'))
            line[--len] = '\0';
        if (len == 0) continue;

        /* Find tabs: chinese\tattempts\tcorrect\tstreak */
        char *t1 = strchr(line, '\t');
        if (!t1) continue;
        *t1 = '\0';
        char *t2 = strchr(t1 + 1, '\t');
        if (!t2) continue;
        *t2 = '\0';
        char *t3 = strchr(t2 + 1, '\t');
        if (!t3) continue;
        *t3 = '\0';

        int attempts = atoi(t1 + 1);
        int correct  = atoi(t2 + 1);
        int streak   = atoi(t3 + 1);

        /* Find matching sentence */
        for (int i = 0; i < ds->num_sentences; i++) {
            if (strcmp(ds->sentences[i].chinese, line) == 0) {
                ds->progress[i].attempts = attempts;
                ds->progress[i].correct  = correct;
                ds->progress[i].streak   = streak;
                break;
            }
        }
    }

    fclose(f);
}

static void drill_save_progress(DrillState *ds, const char *filepath)
{
    /* Ensure parent directory exists */
    char dir[MAX_PATH];
    strncpy(dir, filepath, MAX_PATH - 1);
    dir[MAX_PATH - 1] = '\0';
    char *last_sep = strrchr(dir, '\\');
    if (!last_sep) last_sep = strrchr(dir, '/');
    if (last_sep) {
        *last_sep = '\0';
        CreateDirectoryA(dir, NULL);  /* OK if exists */
    }

    FILE *f = fopen(filepath, "w");
    if (!f) return;

    for (int i = 0; i < ds->num_sentences; i++) {
        if (ds->progress[i].attempts > 0) {
            fprintf(f, "%s\t%d\t%d\t%d\n",
                    ds->sentences[i].chinese,
                    ds->progress[i].attempts,
                    ds->progress[i].correct,
                    ds->progress[i].streak);
        }
    }

    fclose(f);
}

/* ---------- Public API ---------- */

int drill_init(DrillState *ds, const char *sentence_file, const char *progress_file)
{
    memset(ds, 0, sizeof(*ds));
    ds->current_idx = -1;
    ds->hsk_filter = 0;  /* all levels */

    int rc = drill_load_sentences(ds, sentence_file);
    if (rc != 0) return rc;

    if (progress_file)
        drill_load_progress(ds, progress_file);

    return 0;
}

void drill_shutdown(DrillState *ds, const char *progress_file)
{
    if (progress_file)
        drill_save_progress(ds, progress_file);
}

int drill_advance(DrillState *ds)
{
    if (ds->num_sentences == 0) return -1;

    /* Build eligible list based on HSK filter */
    int eligible[DRILL_MAX_SENTENCES];
    double weights[DRILL_MAX_SENTENCES];
    int num_eligible = 0;
    double total_weight = 0.0;

    for (int i = 0; i < ds->num_sentences; i++) {
        if (ds->hsk_filter > 0 && ds->sentences[i].hsk_level != ds->hsk_filter)
            continue;

        eligible[num_eligible] = i;

        /* Weight: inverse of accuracy. Untested sentences get high weight. */
        double accuracy = 0.0;
        if (ds->progress[i].attempts > 0)
            accuracy = (double)ds->progress[i].correct / ds->progress[i].attempts;

        double w = 1.0 / (accuracy + 0.1);

        /* Reduce weight for sentences with long streaks */
        if (ds->progress[i].streak >= 3) w *= 0.3;

        /* Avoid immediate repeat */
        if (i == ds->current_idx) w *= 0.01;

        weights[num_eligible] = w;
        total_weight += w;
        num_eligible++;
    }

    if (num_eligible == 0) return -1;

    /* Weighted random selection */
    double r = ((double)rand() / RAND_MAX) * total_weight;
    double cumulative = 0.0;
    int selected = eligible[0];
    for (int i = 0; i < num_eligible; i++) {
        cumulative += weights[i];
        if (r <= cumulative) {
            selected = eligible[i];
            break;
        }
    }

    ds->current_idx = selected;
    ds->has_result = 0;
    memset(&ds->last_diff, 0, sizeof(ds->last_diff));
    ds->result_text[0] = '\0';

    return selected;
}

int drill_check(DrillState *ds, const char *actual)
{
    if (ds->current_idx < 0 || ds->current_idx >= ds->num_sentences) return 0;

    DrillDiff *d = &ds->last_diff;
    memset(d, 0, sizeof(*d));

    const char *expected = ds->sentences[ds->current_idx].chinese;

    /* Extract codepoints */
    d->num_expected = utf8_to_codepoints(expected, d->expected_cps, DRILL_MAX_TEXT);
    d->num_actual   = utf8_to_codepoints(actual,   d->actual_cps,   DRILL_MAX_TEXT);

    /* Strip punctuation from actual (ASR output) */
    d->num_actual = strip_codepoints(d->actual_cps, d->num_actual);

    /* Position-based comparison */
    int min_len = d->num_expected < d->num_actual ? d->num_expected : d->num_actual;
    int all_match = 1;

    for (int i = 0; i < min_len; i++) {
        d->char_match[i] = homophones_match(d->expected_cps[i], d->actual_cps[i]);
        if (!d->char_match[i]) all_match = 0;
    }

    /* Length mismatch means not a perfect match */
    if (d->num_expected != d->num_actual) all_match = 0;

    /* Mark extra positions as mismatches */
    for (int i = min_len; i < d->num_expected; i++)
        d->char_match[i] = 0;

    d->match = all_match;

    /* Store result text */
    strncpy(ds->result_text, actual, DRILL_MAX_TEXT - 1);
    ds->result_text[DRILL_MAX_TEXT - 1] = '\0';
    ds->has_result = 1;

    return all_match;
}

void drill_record_attempt(DrillState *ds, int correct)
{
    if (ds->current_idx < 0 || ds->current_idx >= ds->num_sentences) return;

    DrillSentenceProgress *p = &ds->progress[ds->current_idx];
    p->attempts++;
    if (correct) {
        p->correct++;
        p->streak++;
    } else {
        p->streak = 0;
    }

    ds->session_attempts++;
    if (correct) ds->session_correct++;
}

/*
 * drill.h -- Pronunciation drill mode for voice note GUI
 *
 * Character-by-character Chinese pronunciation trainer.
 * The ASR model is the judge: user speaks a target sentence,
 * ASR output is diffed against the target.
 */

#ifndef DRILL_H
#define DRILL_H

#define DRILL_MAX_SENTENCES 500
#define DRILL_MAX_TEXT      256

typedef struct {
    char chinese[DRILL_MAX_TEXT];   /* UTF-8 Chinese characters */
    char pinyin[DRILL_MAX_TEXT];    /* Tone numbers: ni3 hao3 */
    char english[DRILL_MAX_TEXT];   /* Translation */
    int  hsk_level;                 /* 1-3 */
} DrillSentence;

typedef struct {
    int attempts;
    int correct;
    int streak;                     /* Consecutive correct */
} DrillSentenceProgress;

typedef struct {
    int match;                                /* 1 if perfect, 0 otherwise */
    int expected_cps[DRILL_MAX_TEXT];          /* Unicode codepoints */
    int actual_cps[DRILL_MAX_TEXT];
    int char_match[DRILL_MAX_TEXT];            /* Per-position: 1=match, 0=miss */
    int num_expected;
    int num_actual;
} DrillDiff;

typedef struct {
    DrillSentence         sentences[DRILL_MAX_SENTENCES];
    DrillSentenceProgress progress[DRILL_MAX_SENTENCES];
    int  num_sentences;
    int  current_idx;
    int  session_attempts;
    int  session_correct;
    DrillDiff last_diff;
    char result_text[DRILL_MAX_TEXT];   /* Last ASR result */
    int  has_result;                    /* 1 after first attempt on current sentence */
    int  hsk_filter;                    /* 0=all, 1-3=specific level */
} DrillState;

/* Load sentence bank from pipe-delimited file. Returns 0 on success. */
int  drill_init(DrillState *ds, const char *sentence_file, const char *progress_file);

/* Save progress and clean up. */
void drill_shutdown(DrillState *ds, const char *progress_file);

/* Select next sentence using weighted selection. Returns sentence index. */
int  drill_advance(DrillState *ds);

/* Compare ASR result against current target. Fills last_diff. Returns 1 if match. */
int  drill_check(DrillState *ds, const char *actual);

/* Record attempt result (call after drill_check). */
void drill_record_attempt(DrillState *ds, int correct);

#endif /* DRILL_H */

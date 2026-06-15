#include "whisper.h"

#if defined(_WIN32)
#define SPELL_EXPORT extern "C" __declspec(dllexport)
#else
#define SPELL_EXPORT extern "C" __attribute__((visibility("default")))
#endif

SPELL_EXPORT whisper_context * spell_whisper_init_from_file(const char * model_path) {
    return whisper_init_from_file(model_path);
}

SPELL_EXPORT void spell_whisper_free(whisper_context * ctx) {
    whisper_free(ctx);
}

SPELL_EXPORT int spell_whisper_full(
    whisper_context * ctx,
    const float * samples,
    int n_samples,
    const char * language
) {
    whisper_full_params params = whisper_full_default_params(WHISPER_SAMPLING_GREEDY);
    params.n_threads = 4;
    params.print_special = false;
    params.print_progress = false;
    params.print_realtime = false;
    params.print_timestamps = false;
    params.language = language;
    return whisper_full(ctx, params, samples, n_samples);
}

SPELL_EXPORT int spell_whisper_full_n_segments(whisper_context * ctx) {
    return whisper_full_n_segments(ctx);
}

SPELL_EXPORT const char * spell_whisper_full_get_segment_text(whisper_context * ctx, int index) {
    return whisper_full_get_segment_text(ctx, index);
}

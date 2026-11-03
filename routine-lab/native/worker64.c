/* Local, disposable Windows x64 machine-code routine worker.
 * Build without -mavx/-march=native: feature detection must run on baseline x64.
 * This process isolates accidental faults; it is not a security sandbox.
 */
#define _WIN32_WINNT 0x0601
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <cpuid.h>
#include <ctype.h>
#include <inttypes.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>

#if !defined(__x86_64__)
#error This worker must be compiled for Windows x64.
#endif

#define MAX_TEXT (256u * 1024u)
#define MAX_CODE (64u * 1024u)
#define SCRATCH_SIZE 4096u

typedef uint64_t (*Routine)(uintptr_t, void *, uint64_t, uintptr_t);
static HANDLE result_file = INVALID_HANDLE_VALUE;
static uint64_t iteration_count;
static int have_avx, have_fma;
static volatile LONG executing;
static LARGE_INTEGER frequency, started;
/* Static storage minimizes the handler's use of a potentially damaged stack. */
static char result_text[1024];
static char exception_text[96];

static int write_result(const char *status, const char *error, double elapsed,
                        uint64_t value, const unsigned char *scratch)
{
    static const char digits[] = "0123456789ABCDEF";
    char hex[33];
    DWORD written;
    int length;
    unsigned i;
    for (i = 0; i < 16; ++i) {
        unsigned byte = scratch ? scratch[i] : 0;
        hex[i * 2] = digits[byte >> 4];
        hex[i * 2 + 1] = digits[byte & 15];
    }
    hex[32] = 0;
    length = snprintf(result_text, sizeof result_text,
        "status=%s\r\narchitecture=x64\r\niterations=%" PRIu64
        "\r\nelapsed_ms=%.6f\r\nreturn_hex=%016" PRIX64
        "\r\nscratch_hex=%s\r\navx=%d\r\nfma=%d\r\n%s%s%s",
        status, iteration_count, elapsed, value, hex, have_avx, have_fma,
        error ? "error=" : "", error ? error : "", error ? "\r\n" : "");
    if (length < 0 || (size_t)length >= sizeof result_text) return 0;
    return WriteFile(result_file, result_text, (DWORD)length, &written, NULL)
        && written == (DWORD)length;
}

static double elapsed_ms(void)
{
    LARGE_INTEGER ended;
    if (!frequency.QuadPart || !QueryPerformanceCounter(&ended)) return 0;
    return (double)(ended.QuadPart - started.QuadPart) * 1000.0
        / (double)frequency.QuadPart;
}

static LONG CALLBACK routine_exception(EXCEPTION_POINTERS *info)
{
    if (!InterlockedExchange(&executing, 0)) return EXCEPTION_CONTINUE_SEARCH;
    snprintf(exception_text, sizeof exception_text,
        "Native exception 0x%08lX; worker terminated.",
        (unsigned long)info->ExceptionRecord->ExceptionCode);
    /* Do not dereference scratch or try to unwind arbitrary generated code. */
    write_result("exception", exception_text, elapsed_ms(), 0, NULL);
    FlushFileBuffers(result_file);
    ExitProcess(4);
    return EXCEPTION_CONTINUE_SEARCH;
}

static void detect_features(void)
{
    unsigned eax, ebx, ecx, edx;
    if (!__get_cpuid(1, &eax, &ebx, &ecx, &edx)) return;
    /* XGETBV itself faults without OSXSAVE. Check all three bits first. */
    if ((ecx & ((1u << 26) | (1u << 27) | (1u << 28)))
             == ((1u << 26) | (1u << 27) | (1u << 28))) {
        unsigned low, high;
        __asm__ volatile ("xgetbv" : "=a" (low), "=d" (high) : "c" (0));
        (void)high;
        have_avx = (low & 6u) == 6u;
        have_fma = have_avx && (ecx & (1u << 12)) != 0;
    }
}

static int parse_iterations(const char *text)
{
    uint64_t value = 0;
    if (!*text) return 0;
    for (; *text; ++text) {
        if (*text < '0' || *text > '9') return 0;
        value = value * 10u + (unsigned)(*text - '0');
        if (value > 2147483647u) return 0;
    }
    if (!value) return 0;
    iteration_count = value;
    return 1;
}

static int hex_digit(unsigned char c)
{
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return -1;
}

static const char *read_code(const char *path, unsigned char *code, size_t *size)
{
    HANDLE file;
    LARGE_INTEGER length;
    DWORD got;
    unsigned char *text;
    size_t pos = 0, total;
    const char *error = NULL;
    *size = 0;
    file = CreateFileA(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING,
                       FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) return "Cannot open hex input file.";
    if (!GetFileSizeEx(file, &length) || length.QuadPart < 0
            || length.QuadPart > MAX_TEXT) {
        CloseHandle(file);
        return "Cannot read input size, or input exceeds 256 KiB.";
    }
    total = (size_t)length.QuadPart;
    text = HeapAlloc(GetProcessHeap(), 0, total + 1);
    if (!text) {
        CloseHandle(file);
        return "Cannot allocate input buffer.";
    }
    if (!ReadFile(file, text, (DWORD)total, &got, NULL) || got != total) {
        error = "Cannot read complete hex input file.";
    }
    CloseHandle(file);
    while (!error && pos < total) {
        int high, low;
        if (isspace(text[pos])) { ++pos; continue; }
        if (text[pos] == ';') {
            while (pos < total && text[pos] != '\n' && text[pos] != '\r') ++pos;
            continue;
        }
        high = hex_digit(text[pos]);
        low = pos + 1 < total ? hex_digit(text[pos + 1]) : -1;
        if (high < 0 || low < 0 || (pos + 2 < total
                && !isspace(text[pos + 2]) && text[pos + 2] != ';')) {
            error = "Expected whitespace-separated two-digit hex bytes.";
            break;
        }
        if (*size == MAX_CODE) {
            error = "Machine code exceeds 64 KiB.";
            break;
        }
        code[(*size)++] = (unsigned char)((high << 4) | low);
        pos += 2;
    }
    HeapFree(GetProcessHeap(), 0, text);
    if (!error && !*size) error = "Hex input contains no machine code.";
    return error;
}

int main(int argc, char **argv)
{
    static unsigned char bytes[MAX_CODE];
    size_t code_size = 0;
    unsigned char *scratch = NULL;
    void *code = NULL;
    PVOID handler = NULL;
    DWORD previous;
    Routine routine;
    uint64_t value;
    double elapsed;
    const char *error = NULL;
    const char *status = "error";
    int exit_code = 2;

    SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOGPFAULTERRORBOX | SEM_NOOPENFILEERRORBOX);
    if (argc != 5) {
        fputs("Usage: RoutineWorker64.exe <hex-file> <iterations> <result-file> <base|sse2|avx|fma>\n", stderr);
        return 2;
    }
    detect_features();
    if (!parse_iterations(argv[2])) error = "Iterations must be decimal 1..2147483647.";
    else if (strcmp(argv[4], "base") && strcmp(argv[4], "sse2")
            && strcmp(argv[4], "avx") && strcmp(argv[4], "fma"))
        error = "Requirement must be base, sse2, avx, or fma.";
    else if ((!strcmp(argv[4], "avx") && !have_avx)
            || (!strcmp(argv[4], "fma") && !have_fma)) {
        error = "Requested instruction set is unavailable in CPU/OS state.";
        status = "unsupported";
        exit_code = 3;
    } else error = read_code(argv[1], bytes, &code_size);

    /* Input parsing completes before the result file is created. */
    result_file = CreateFileA(argv[3], GENERIC_WRITE, FILE_SHARE_READ, NULL,
                             CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (result_file == INVALID_HANDLE_VALUE) {
        fputs("Cannot create result file.\n", stderr);
        return 2;
    }
    if (error) goto failed;
    code = VirtualAlloc(NULL, code_size, MEM_RESERVE | MEM_COMMIT, PAGE_READWRITE);
    scratch = VirtualAlloc(NULL, SCRATCH_SIZE, MEM_RESERVE | MEM_COMMIT, PAGE_READWRITE);
    if (!code || !scratch) { error = "VirtualAlloc failed."; goto failed; }
    memcpy(code, bytes, code_size);
    if (!VirtualProtect(code, code_size, PAGE_EXECUTE_READ, &previous)) {
        error = "VirtualProtect could not make code executable."; goto failed;
    }
    if (!FlushInstructionCache(GetCurrentProcess(), code, code_size)) {
        error = "FlushInstructionCache failed."; goto failed;
    }
    if (!QueryPerformanceFrequency(&frequency) || frequency.QuadPart <= 0) {
        error = "Performance counter is unavailable."; goto failed;
    }
    handler = AddVectoredExceptionHandler(1, routine_exception);
    if (!handler) { error = "Cannot register exception handler."; goto failed; }
    /* Windows x64 function and object pointers have identical representations. */
    _Static_assert(sizeof routine == sizeof code, "Unexpected pointer representation");
    memcpy(&routine, &code, sizeof routine);
    InterlockedExchange(&executing, 1);
    if (!QueryPerformanceCounter(&started)) {
        InterlockedExchange(&executing, 0);
        error = "Cannot start performance counter."; goto failed;
    }
    value = routine(0, scratch, iteration_count, SCRATCH_SIZE);
    elapsed = elapsed_ms();
    InterlockedExchange(&executing, 0);
    exit_code = write_result("ok", NULL, elapsed, value, scratch) ? 0 : 2;
    goto done;

failed:
    write_result(status, error, 0, 0, NULL);
done:
    if (handler) RemoveVectoredExceptionHandler(handler);
    if (code) VirtualFree(code, 0, MEM_RELEASE);
    if (scratch) VirtualFree(scratch, 0, MEM_RELEASE);
    CloseHandle(result_file);
    return exit_code;
}

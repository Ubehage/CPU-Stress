BITS 32
; Legacy fallback: four balanced x87 accumulators plus integer mixing.
; No SSE/AVX/FXSAVE required. Entry/exit x87 stack depth is unchanged.
; x86 stdcall: scratch=[ebp+12], iterations=[ebp+16], RET 16.
; Reset four finite seeds every256 iterations before values can converge.
push ebp
mov ebp, esp
push ebx
push esi
push edi
mov esi, [ebp+16]
xor eax, eax
test esi, esi
jz .empty
mov ebx, [ebp+12]
call .address
.address:
pop ecx
add ecx, data-.address
mov eax, 0x12345678
.reset:
fld dword [ecx+seed0-data]
fld dword [ecx+seed1-data]
fld dword [ecx+seed2-data]
fld dword [ecx+seed3-data]
mov edi, 256
.loop:
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
fxch st1
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
fxch st2
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
fxch st3
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
imul eax, eax, 1664525
add eax, 1013904223
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
fxch st1
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
fxch st2
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
fxch st3
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
imul eax, eax, 1664525
add eax, 1013904223
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
fxch st1
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
fxch st2
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
fxch st3
fmul dword [ecx+factor_down-data]
fadd dword [ecx+bias_up-data]
imul eax, eax, 1664525
add eax, 1013904223
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
fxch st1
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
fxch st2
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
fxch st3
fmul dword [ecx+factor_up-data]
fadd dword [ecx+bias_down-data]
imul eax, eax, 1664525
add eax, 1013904223
dec esi
jz .finished
dec edi
jnz .loop
fstp st0
fstp st0
fstp st0
fstp st0
jmp .reset
.finished:
fstp dword [ebx+0]
fstp dword [ebx+4]
fstp dword [ebx+8]
fstp dword [ebx+12]
mov [ebx+16], eax
imul eax, eax, 1664525
add eax, 1013904223
mov [ebx+20], eax
imul eax, eax, 1664525
add eax, 1013904223
mov [ebx+24], eax
imul eax, eax, 1664525
add eax, 1013904223
mov [ebx+28], eax
mov eax, [ebp+16]
.empty:
pop edi
pop esi
pop ebx
pop ebp
ret 16
align 16, db 0
data:
seed0: dd 0x3e000000
seed1: dd 0x3e800000
seed2: dd 0x3ec00000
seed3: dd 0x3f000000
factor_down: dd 0x3f7fffff
bias_up: dd 0x35800000
factor_up: dd 0x3f800001
bias_down: dd 0xb5000000

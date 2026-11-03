BITS 32
; SSE32: four independent packed-single accumulators; 128 FLOPs/iteration.
; x86 stdcall: [ebp+12]=scratch, [ebp+16]=iterations; RET 16.
; Same alternating finite recurrence and 256-iteration resets as x64 kernels.
; Preserve nonvolatile GPRs; XMM registers are volatile in the x86 ABI.
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
movups xmm4, [ecx+factor_down-data]
movups xmm5, [ecx+bias_up-data]
movups xmm6, [ecx+factor_up-data]
movups xmm7, [ecx+bias_down-data]
.reset:
movups xmm0, [ecx+seed0-data]
movups xmm1, [ecx+seed1-data]
movups xmm2, [ecx+seed2-data]
movups xmm3, [ecx+seed3-data]
mov edi, 256
.loop:
mulps xmm0, xmm4
mulps xmm1, xmm4
mulps xmm2, xmm4
mulps xmm3, xmm4
addps xmm0, xmm5
addps xmm1, xmm5
addps xmm2, xmm5
addps xmm3, xmm5
mulps xmm0, xmm6
mulps xmm1, xmm6
mulps xmm2, xmm6
mulps xmm3, xmm6
addps xmm0, xmm7
addps xmm1, xmm7
addps xmm2, xmm7
addps xmm3, xmm7
mulps xmm0, xmm4
mulps xmm1, xmm4
mulps xmm2, xmm4
mulps xmm3, xmm4
addps xmm0, xmm5
addps xmm1, xmm5
addps xmm2, xmm5
addps xmm3, xmm5
mulps xmm0, xmm6
mulps xmm1, xmm6
mulps xmm2, xmm6
mulps xmm3, xmm6
addps xmm0, xmm7
addps xmm1, xmm7
addps xmm2, xmm7
addps xmm3, xmm7
dec esi
jz .finished
dec edi
jnz .loop
jmp .reset
.finished:
movups [ebx+0], xmm0
movups [ebx+16], xmm1
movups [ebx+32], xmm2
movups [ebx+48], xmm3
mov eax, [ebp+16]
.empty:
pop edi
pop esi
pop ebx
pop ebp
ret 16
align 16, db 0
data:
seed0: dd 0x3e000000, 0x3e100000, 0x3e200000, 0x3e300000
seed1: dd 0x3e800000, 0x3e880000, 0x3e900000, 0x3e980000
seed2: dd 0x3ec00000, 0x3ec80000, 0x3ed00000, 0x3ed80000
seed3: dd 0x3f000000, 0x3f040000, 0x3f080000, 0x3f0c0000
factor_down: dd 0x3f7fffff, 0x3f7fffff, 0x3f7fffff, 0x3f7fffff
bias_up: dd 0x35800000, 0x35800000, 0x35800000, 0x35800000
factor_up: dd 0x3f800001, 0x3f800001, 0x3f800001, 0x3f800001
bias_down: dd 0xb5000000, 0xb5000000, 0xb5000000, 0xb5000000

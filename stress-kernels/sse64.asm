BITS 64
; SSE64: 8 independent 128-bit accumulators; four updates per iteration.
; 256 floating-point operations per iteration. No Windows API calls.
; Windows x64 ABI: RDX=scratch, R8=iterations, R9=usable bytes.
; Every 256 iterations restart different finite seeds, avoiding convergence.
; Updates alternate x*(1-2^-24)+2^-20 and x*(1+2^-23)-2^-21.
; Store every accumulator; return the requested iteration count.
mov r10, r8
xor eax, eax
test r10, r10
jz .empty
sub rsp, 104
movdqa [rsp+0], xmm6
movdqa [rsp+16], xmm7
movdqa [rsp+32], xmm8
movdqa [rsp+48], xmm9
movdqa [rsp+64], xmm10
movdqa [rsp+80], xmm11
movups xmm8, [rel factor_down]
movups xmm9, [rel bias_up]
movups xmm10, [rel factor_up]
movups xmm11, [rel bias_down]
.reset:
movups xmm0, [rel seed0]
movups xmm1, [rel seed1]
movups xmm2, [rel seed2]
movups xmm3, [rel seed3]
movups xmm4, [rel seed4]
movups xmm5, [rel seed5]
movups xmm6, [rel seed6]
movups xmm7, [rel seed7]
mov r11d, 256
.loop:
; Update 1: slightly below one, positive bias.
mulps xmm0, xmm8
mulps xmm1, xmm8
mulps xmm2, xmm8
mulps xmm3, xmm8
mulps xmm4, xmm8
mulps xmm5, xmm8
mulps xmm6, xmm8
mulps xmm7, xmm8
addps xmm0, xmm9
addps xmm1, xmm9
addps xmm2, xmm9
addps xmm3, xmm9
addps xmm4, xmm9
addps xmm5, xmm9
addps xmm6, xmm9
addps xmm7, xmm9
; Update 2: slightly above one, negative bias.
mulps xmm0, xmm10
mulps xmm1, xmm10
mulps xmm2, xmm10
mulps xmm3, xmm10
mulps xmm4, xmm10
mulps xmm5, xmm10
mulps xmm6, xmm10
mulps xmm7, xmm10
addps xmm0, xmm11
addps xmm1, xmm11
addps xmm2, xmm11
addps xmm3, xmm11
addps xmm4, xmm11
addps xmm5, xmm11
addps xmm6, xmm11
addps xmm7, xmm11
; Update 3: slightly below one, positive bias.
mulps xmm0, xmm8
mulps xmm1, xmm8
mulps xmm2, xmm8
mulps xmm3, xmm8
mulps xmm4, xmm8
mulps xmm5, xmm8
mulps xmm6, xmm8
mulps xmm7, xmm8
addps xmm0, xmm9
addps xmm1, xmm9
addps xmm2, xmm9
addps xmm3, xmm9
addps xmm4, xmm9
addps xmm5, xmm9
addps xmm6, xmm9
addps xmm7, xmm9
; Update 4: slightly above one, negative bias.
mulps xmm0, xmm10
mulps xmm1, xmm10
mulps xmm2, xmm10
mulps xmm3, xmm10
mulps xmm4, xmm10
mulps xmm5, xmm10
mulps xmm6, xmm10
mulps xmm7, xmm10
addps xmm0, xmm11
addps xmm1, xmm11
addps xmm2, xmm11
addps xmm3, xmm11
addps xmm4, xmm11
addps xmm5, xmm11
addps xmm6, xmm11
addps xmm7, xmm11
dec r10
jz .finished
dec r11d
jnz .loop
jmp .reset
.finished:
movups [rdx+0], xmm0
movups [rdx+16], xmm1
movups [rdx+32], xmm2
movups [rdx+48], xmm3
movups [rdx+64], xmm4
movups [rdx+80], xmm5
movups [rdx+96], xmm6
movups [rdx+112], xmm7
movdqa xmm6, [rsp+0]
movdqa xmm7, [rsp+16]
movdqa xmm8, [rsp+32]
movdqa xmm9, [rsp+48]
movdqa xmm10, [rsp+64]
movdqa xmm11, [rsp+80]
add rsp, 104
mov rax, r8
.empty:
ret
align 16, db 0
data:
seed0: dd 0x3e000000, 0x3e100000, 0x3e200000, 0x3e300000
seed1: dd 0x3e800000, 0x3e880000, 0x3e900000, 0x3e980000
seed2: dd 0x3ec00000, 0x3ec80000, 0x3ed00000, 0x3ed80000
seed3: dd 0x3f000000, 0x3f040000, 0x3f080000, 0x3f0c0000
seed4: dd 0x3f200000, 0x3f240000, 0x3f280000, 0x3f2c0000
seed5: dd 0x3f400000, 0x3f440000, 0x3f480000, 0x3f4c0000
seed6: dd 0x3f600000, 0x3f640000, 0x3f680000, 0x3f6c0000
seed7: dd 0x3f800000, 0x3f820000, 0x3f840000, 0x3f860000
factor_down: dd 0x3f7fffff, 0x3f7fffff, 0x3f7fffff, 0x3f7fffff
bias_up: dd 0x35800000, 0x35800000, 0x35800000, 0x35800000
factor_up: dd 0x3f800001, 0x3f800001, 0x3f800001, 0x3f800001
bias_down: dd 0xb5000000, 0xb5000000, 0xb5000000, 0xb5000000

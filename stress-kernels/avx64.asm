BITS 64
; AVX64: 12 independent 256-bit accumulators; four updates per iteration.
; 768 floating-point operations per iteration. No Windows API calls.
; Windows x64 ABI: RDX=scratch, R8=iterations, R9=usable bytes.
; Every 256 iterations restart different finite seeds, avoiding convergence.
; Updates alternate x*(1-2^-24)+2^-20 and x*(1+2^-23)-2^-21.
; Store every accumulator; return the requested iteration count.
mov r10, r8
xor eax, eax
test r10, r10
jz .empty
sub rsp, 168
movdqa [rsp+0], xmm6
movdqa [rsp+16], xmm7
movdqa [rsp+32], xmm8
movdqa [rsp+48], xmm9
movdqa [rsp+64], xmm10
movdqa [rsp+80], xmm11
movdqa [rsp+96], xmm12
movdqa [rsp+112], xmm13
movdqa [rsp+128], xmm14
movdqa [rsp+144], xmm15
vmovups ymm12, [rel factor_down]
vmovups ymm13, [rel bias_up]
vmovups ymm14, [rel factor_up]
vmovups ymm15, [rel bias_down]
.reset:
vmovups ymm0, [rel seed0]
vmovups ymm1, [rel seed1]
vmovups ymm2, [rel seed2]
vmovups ymm3, [rel seed3]
vmovups ymm4, [rel seed4]
vmovups ymm5, [rel seed5]
vmovups ymm6, [rel seed6]
vmovups ymm7, [rel seed7]
vmovups ymm8, [rel seed8]
vmovups ymm9, [rel seed9]
vmovups ymm10, [rel seed10]
vmovups ymm11, [rel seed11]
mov r11d, 256
.loop:
; Update 1: slightly below one, positive bias.
vmulps ymm0, ymm0, ymm12
vmulps ymm1, ymm1, ymm12
vmulps ymm2, ymm2, ymm12
vmulps ymm3, ymm3, ymm12
vmulps ymm4, ymm4, ymm12
vmulps ymm5, ymm5, ymm12
vmulps ymm6, ymm6, ymm12
vmulps ymm7, ymm7, ymm12
vmulps ymm8, ymm8, ymm12
vmulps ymm9, ymm9, ymm12
vmulps ymm10, ymm10, ymm12
vmulps ymm11, ymm11, ymm12
vaddps ymm0, ymm0, ymm13
vaddps ymm1, ymm1, ymm13
vaddps ymm2, ymm2, ymm13
vaddps ymm3, ymm3, ymm13
vaddps ymm4, ymm4, ymm13
vaddps ymm5, ymm5, ymm13
vaddps ymm6, ymm6, ymm13
vaddps ymm7, ymm7, ymm13
vaddps ymm8, ymm8, ymm13
vaddps ymm9, ymm9, ymm13
vaddps ymm10, ymm10, ymm13
vaddps ymm11, ymm11, ymm13
; Update 2: slightly above one, negative bias.
vmulps ymm0, ymm0, ymm14
vmulps ymm1, ymm1, ymm14
vmulps ymm2, ymm2, ymm14
vmulps ymm3, ymm3, ymm14
vmulps ymm4, ymm4, ymm14
vmulps ymm5, ymm5, ymm14
vmulps ymm6, ymm6, ymm14
vmulps ymm7, ymm7, ymm14
vmulps ymm8, ymm8, ymm14
vmulps ymm9, ymm9, ymm14
vmulps ymm10, ymm10, ymm14
vmulps ymm11, ymm11, ymm14
vaddps ymm0, ymm0, ymm15
vaddps ymm1, ymm1, ymm15
vaddps ymm2, ymm2, ymm15
vaddps ymm3, ymm3, ymm15
vaddps ymm4, ymm4, ymm15
vaddps ymm5, ymm5, ymm15
vaddps ymm6, ymm6, ymm15
vaddps ymm7, ymm7, ymm15
vaddps ymm8, ymm8, ymm15
vaddps ymm9, ymm9, ymm15
vaddps ymm10, ymm10, ymm15
vaddps ymm11, ymm11, ymm15
; Update 3: slightly below one, positive bias.
vmulps ymm0, ymm0, ymm12
vmulps ymm1, ymm1, ymm12
vmulps ymm2, ymm2, ymm12
vmulps ymm3, ymm3, ymm12
vmulps ymm4, ymm4, ymm12
vmulps ymm5, ymm5, ymm12
vmulps ymm6, ymm6, ymm12
vmulps ymm7, ymm7, ymm12
vmulps ymm8, ymm8, ymm12
vmulps ymm9, ymm9, ymm12
vmulps ymm10, ymm10, ymm12
vmulps ymm11, ymm11, ymm12
vaddps ymm0, ymm0, ymm13
vaddps ymm1, ymm1, ymm13
vaddps ymm2, ymm2, ymm13
vaddps ymm3, ymm3, ymm13
vaddps ymm4, ymm4, ymm13
vaddps ymm5, ymm5, ymm13
vaddps ymm6, ymm6, ymm13
vaddps ymm7, ymm7, ymm13
vaddps ymm8, ymm8, ymm13
vaddps ymm9, ymm9, ymm13
vaddps ymm10, ymm10, ymm13
vaddps ymm11, ymm11, ymm13
; Update 4: slightly above one, negative bias.
vmulps ymm0, ymm0, ymm14
vmulps ymm1, ymm1, ymm14
vmulps ymm2, ymm2, ymm14
vmulps ymm3, ymm3, ymm14
vmulps ymm4, ymm4, ymm14
vmulps ymm5, ymm5, ymm14
vmulps ymm6, ymm6, ymm14
vmulps ymm7, ymm7, ymm14
vmulps ymm8, ymm8, ymm14
vmulps ymm9, ymm9, ymm14
vmulps ymm10, ymm10, ymm14
vmulps ymm11, ymm11, ymm14
vaddps ymm0, ymm0, ymm15
vaddps ymm1, ymm1, ymm15
vaddps ymm2, ymm2, ymm15
vaddps ymm3, ymm3, ymm15
vaddps ymm4, ymm4, ymm15
vaddps ymm5, ymm5, ymm15
vaddps ymm6, ymm6, ymm15
vaddps ymm7, ymm7, ymm15
vaddps ymm8, ymm8, ymm15
vaddps ymm9, ymm9, ymm15
vaddps ymm10, ymm10, ymm15
vaddps ymm11, ymm11, ymm15
dec r10
jz .finished
dec r11d
jnz .loop
jmp .reset
.finished:
vmovups [rdx+0], ymm0
vmovups [rdx+32], ymm1
vmovups [rdx+64], ymm2
vmovups [rdx+96], ymm3
vmovups [rdx+128], ymm4
vmovups [rdx+160], ymm5
vmovups [rdx+192], ymm6
vmovups [rdx+224], ymm7
vmovups [rdx+256], ymm8
vmovups [rdx+288], ymm9
vmovups [rdx+320], ymm10
vmovups [rdx+352], ymm11
vzeroupper
movdqa xmm6, [rsp+0]
movdqa xmm7, [rsp+16]
movdqa xmm8, [rsp+32]
movdqa xmm9, [rsp+48]
movdqa xmm10, [rsp+64]
movdqa xmm11, [rsp+80]
movdqa xmm12, [rsp+96]
movdqa xmm13, [rsp+112]
movdqa xmm14, [rsp+128]
movdqa xmm15, [rsp+144]
add rsp, 168
mov rax, r8
.empty:
ret
align 16, db 0
data:
seed0: dd 0x3e000000, 0x3e100000, 0x3e200000, 0x3e300000, 0x3e400000, 0x3e500000, 0x3e600000, 0x3e700000
seed1: dd 0x3e800000, 0x3e880000, 0x3e900000, 0x3e980000, 0x3ea00000, 0x3ea80000, 0x3eb00000, 0x3eb80000
seed2: dd 0x3ec00000, 0x3ec80000, 0x3ed00000, 0x3ed80000, 0x3ee00000, 0x3ee80000, 0x3ef00000, 0x3ef80000
seed3: dd 0x3f000000, 0x3f040000, 0x3f080000, 0x3f0c0000, 0x3f100000, 0x3f140000, 0x3f180000, 0x3f1c0000
seed4: dd 0x3f200000, 0x3f240000, 0x3f280000, 0x3f2c0000, 0x3f300000, 0x3f340000, 0x3f380000, 0x3f3c0000
seed5: dd 0x3f400000, 0x3f440000, 0x3f480000, 0x3f4c0000, 0x3f500000, 0x3f540000, 0x3f580000, 0x3f5c0000
seed6: dd 0x3f600000, 0x3f640000, 0x3f680000, 0x3f6c0000, 0x3f700000, 0x3f740000, 0x3f780000, 0x3f7c0000
seed7: dd 0x3f800000, 0x3f820000, 0x3f840000, 0x3f860000, 0x3f880000, 0x3f8a0000, 0x3f8c0000, 0x3f8e0000
seed8: dd 0x3f900000, 0x3f920000, 0x3f940000, 0x3f960000, 0x3f980000, 0x3f9a0000, 0x3f9c0000, 0x3f9e0000
seed9: dd 0x3fa00000, 0x3fa20000, 0x3fa40000, 0x3fa60000, 0x3fa80000, 0x3faa0000, 0x3fac0000, 0x3fae0000
seed10: dd 0x3fb00000, 0x3fb20000, 0x3fb40000, 0x3fb60000, 0x3fb80000, 0x3fba0000, 0x3fbc0000, 0x3fbe0000
seed11: dd 0x3fc00000, 0x3fc20000, 0x3fc40000, 0x3fc60000, 0x3fc80000, 0x3fca0000, 0x3fcc0000, 0x3fce0000
factor_down: dd 0x3f7fffff, 0x3f7fffff, 0x3f7fffff, 0x3f7fffff, 0x3f7fffff, 0x3f7fffff, 0x3f7fffff, 0x3f7fffff
bias_up: dd 0x35800000, 0x35800000, 0x35800000, 0x35800000, 0x35800000, 0x35800000, 0x35800000, 0x35800000
factor_up: dd 0x3f800001, 0x3f800001, 0x3f800001, 0x3f800001, 0x3f800001, 0x3f800001, 0x3f800001, 0x3f800001
bias_down: dd 0xb5000000, 0xb5000000, 0xb5000000, 0xb5000000, 0xb5000000, 0xb5000000, 0xb5000000, 0xb5000000

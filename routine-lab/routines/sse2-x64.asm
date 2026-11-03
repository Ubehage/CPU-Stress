BITS 64
; Same four-chain finite packed-single recurrence as sse2-x86.asm.
; 32 FLOPs/iteration. Only volatile XMM0..5 and volatile GPRs are changed.
mov r10, r8
xor eax, eax
test r10, r10
jz .done
mov eax, 0x3f000000
movd xmm0, eax
shufps xmm0, xmm0, 0
movaps xmm1, xmm0
movaps xmm2, xmm0
movaps xmm3, xmm0
mov eax, 0x3f7fffff
movd xmm4, eax
shufps xmm4, xmm4, 0
mov eax, 0x34000000
movd xmm5, eax
shufps xmm5, xmm5, 0
.loop:
mulps xmm0, xmm4
mulps xmm1, xmm4
mulps xmm2, xmm4
mulps xmm3, xmm4
addps xmm0, xmm5
addps xmm1, xmm5
addps xmm2, xmm5
addps xmm3, xmm5
dec r10
jnz .loop
movups [rdx], xmm0
mov rax, r8
.done:
ret

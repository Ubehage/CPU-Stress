BITS 32
; Four independent packed-single recurrence chains, 32 FLOPs/iteration.
; x = x * (1 - 2^-24) + 2^-23, initially 0.5 in every lane.
; The finite recurrence avoids overflow, underflow, and denormals.
; ABI: [esp+8]=scratch, [esp+12]=iterations; preserve nonvolatile GPRs.
mov ecx, [esp+12]
xor eax, eax
test ecx, ecx
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
dec ecx
jnz .loop
mov edx, [esp+8]
movups [edx], xmm0
mov eax, [esp+12]
.done:
ret 16

BITS 64
; Eight independent 256-bit FMA chains: 128 FLOPs/iteration.
; x = x * (1 - 2^-24) + 2^-23, initially 0.5 in every lane.
; Windows x64 ABI requires preserving the low halves of XMM6..15.
; The incoming stack is 8 mod 16; subtract 88 to align these saves.
mov r10, r8
xor eax, eax
test r10, r10
jz .done
sub rsp, 88
movdqa [rsp], xmm6
movdqa [rsp+16], xmm7
movdqa [rsp+32], xmm8
movdqa [rsp+48], xmm9
mov dword [rsp+64], 0x3f7fffff
mov dword [rsp+68], 0x34000000
mov dword [rsp+72], 0x3f000000
vbroadcastss ymm8, [rsp+64]
vbroadcastss ymm9, [rsp+68]
vbroadcastss ymm0, [rsp+72]
vmovaps ymm1, ymm0
vmovaps ymm2, ymm0
vmovaps ymm3, ymm0
vmovaps ymm4, ymm0
vmovaps ymm5, ymm0
vmovaps ymm6, ymm0
vmovaps ymm7, ymm0
.loop:
vfmadd213ps ymm0, ymm8, ymm9
vfmadd213ps ymm1, ymm8, ymm9
vfmadd213ps ymm2, ymm8, ymm9
vfmadd213ps ymm3, ymm8, ymm9
vfmadd213ps ymm4, ymm8, ymm9
vfmadd213ps ymm5, ymm8, ymm9
vfmadd213ps ymm6, ymm8, ymm9
vfmadd213ps ymm7, ymm8, ymm9
dec r10
jnz .loop
vmovups [rdx], xmm0
vzeroupper
movdqa xmm6, [rsp]
movdqa xmm7, [rsp+16]
movdqa xmm8, [rsp+32]
movdqa xmm9, [rsp+48]
add rsp, 88
mov rax, r8
.done:
ret

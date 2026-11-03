BITS 64
; Windows x64 ABI: RCX=0, RDX=scratch, R8=iterations, R9=scratch size.
; Exercise extended registers, a 64-bit immediate, arithmetic above 32 bits,
; and a 64-bit store. Result: 1122334555667788 (hex).
xor eax, eax
test r8, r8
jz .done
mov r10, 0x1122334455667788
mov r11, 0x100000000
add r10, r11
mov [rdx], r10
mov rax, r10
.done:
ret

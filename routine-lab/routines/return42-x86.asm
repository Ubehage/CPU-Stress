BITS 32
; WindowProc-compatible x86 routine. Four stack arguments, callee cleanup.
; Return a known value without changing the scratch buffer.
xor eax, eax
cmp dword [esp+12], 0
je .done
mov eax, 42
.done:
ret 16

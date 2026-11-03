; Position-independent x86 stdcall -> x64 leaf routine -> x86 bridge.
; Experimental Windows AMD64 WOW64 selectors: CS32=23h, CS64=33h.
; Payload is appended immediately after this binary. No imports or relocations.
BITS 32
    push ebp
    mov ebp, esp
    push ebx
    push esi
    push edi
    xor eax, eax
    mov ax, fs
    push eax
    pushfd
    mov ax, 2bh
    mov fs, ax
    call address32
address32:
    pop eax
    add eax, entry64 - address32
    push 33h
    push eax
    retf

BITS 64
entry64:
    mov esp, esp                 ; explicitly zero-extend the 32-bit stack
    push rbx
    push rsi
    push rdi
    push rbp
    push r12
    push r13
    push r14
    push r15
    mov ebp, ebp
    mov r12, rsp
    and rsp, -16
    sub rsp, 544                 ; 32-byte shadow space + 512-byte FXSAVE area
    fxsave64 [rsp + 32]
    mov ecx, [rbp + 8]
    mov edx, [rbp + 12]
    mov r8d, [rbp + 16]
    mov r9d, [rbp + 20]
    sub r9d, 16                  ; bridge owns the final 16 scratch bytes
    call payload
    mov edx, [rbp + 12]
    mov [rdx + 4080], rax        ; preserve the complete 64-bit result
    mov ax, cs
    mov [rdx + 4088], ax
    fxrstor64 [rsp + 32]
    mov rsp, r12
    pop r15
    pop r14
    pop r13
    pop r12
    pop rbp
    pop rdi
    pop rsi
    pop rbx
    lea rax, [rel return32]
    ; Plain RETF uses a 32-bit offset and selector slot (8 bytes total).
    sub rsp, 8
    mov [rsp], eax
    mov dword [rsp + 4], 23h
    retf

BITS 32
return32:
    mov ax, ds
    mov ss, ax
    popfd
    pop eax
    mov fs, ax
    mov eax, [ebp + 12]
    mov eax, [eax + 4080]
    pop edi
    pop esi
    pop ebx
    pop ebp
    ret 16
payload:

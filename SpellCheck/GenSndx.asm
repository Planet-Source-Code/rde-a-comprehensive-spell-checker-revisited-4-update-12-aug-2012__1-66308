;
 ;   // Russell Soundex
 ;
 ;   // From Wikipedia, the free encyclopedia
 ;
 ;   // Soundex is a phonetic algorithm for indexing names by their
 ;   // sound when pronounced in English.
 ;
 ;   // Soundex is the most widely known of all phonetic algorithms and
 ;   // is often used (incorrectly) as a synonym for "phonetic algorithm".
 ;
 ;   // The basic aim is for names with the same pronunciation to be
 ;   // encoded to the same signature so that matching can occur despite
 ;   // minor differences in spelling.
 ;
 ;   // The Soundex code for a name consists of a letter followed by three
 ;   // numbers: the letter is the first letter of the name, and the numbers
 ;   // encode the remaining consonants.
 ;
 ;   // Similar sounding consonants share the same number so, for example,
 ;   // the labial B, F, P and V are all encoded as 1.
 ;
 ;   // If two or more letters with the same number were adjacent in the
 ;   // original name, or adjacent except for any intervening vowels, then
 ;   // all are omitted except the first.
 ;
 ;   // Vowels can affect the coding, but are never coded directly unless
 ;   // they appear at the start of the name.
 ;
 ;      The vowels (oral resonants)    a, e, i, o, u, y
 ;      The labials and labio-dentals  b, f, p, v
 ;      The gutterals and sibilants    c, g, k, q, s, x, z
 ;      The dental-mutes               d, t
 ;      The palatal-fricative          l
 ;      The labio-nasal                m
 ;      The den to or lingua-nasal     n
 ;      The dental fricative           r
 ;
 ;   // Russell Soundex for Spell Checking
 ;
 ;   // This particular version of the Soundex algorithm has been adapted
 ;   // from the original design in an attempt to more reliably facilitate
 ;   // word matching for a generic English language spell checker.
 ;
 ;   // Normally, each Soundex begins with the first letter of the given
 ;   // name and only subsequent letters are used to produce the phonetic
 ;   // signature, so only names beginning with the same first letter are
 ;   // compared for similar pronunciation using the standard algorithm.
 ;
 ;   // For example, one may seek the correct spelling for "upholstery" and
 ;   // may inadvertently type "apolstry", "apolstery", or even "apholstery"
 ;   // but would still not retrieve the correct spelling for this word.
 ;
 ;   // Therefore, this version of the Soundex algorithm has been modified
 ;   // to allow the matching of words that start with differing first
 ;   // letters so as not to assume that the first letter is always known.
 ;
 ;   // Consequently, encoding begins with the first letter of the word.
 ;
 ;   // Because of this change, many more similarly spelled words are
 ;   // returned as a match, so the Soundex's length has also been
 ;   // extended from three numbers to four to produce a more unique
 ;   // phonetic signature.
;

.486             ; 32-bit instruction set
.model small     ; generate segment ASSUMEs
.code            ; start of code

genSndx proc near   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

    push ebp              ; preserve base pointer value
    mov ebp, esp          ; set up stack pointer to args
    sub esp, 8            ; offset stack pointer past vars

    increm  equ [ebp-8]   ; set up stack variables
    cTotal  equ [ebp-4]

    aWords  equ [ebp+8]   ; pointer to words array
    lA_sdx  equ [ebp+12]  ; pointer to soundex results
    lA_rev  equ [ebp+16]  ; pointer to reverse soundex
    lA_len  equ [ebp+20]  ; pointer to str lengths array

    pushf        ; preserve the flags register
    push esi     ; preserve registers before use
    push edi
    push edx
    push ecx     ; convention does not require ecx
    push ebx

    xor eax, eax
    mov increm, eax    ; set incrementer to zero
    mov edi, lA_len    ; set pointer to length array
    mov ebx, [edi]     ; copy out words array length
    mov cTotal, ebx    ; preserve the word count

  _wordsLoop:
    cmp eax, cTotal
    je _exitp
    shl eax, 2         ; mul 4 bytes per descriptor
    mov edi, aWords    ; copy str descriptor to edi
    add edi, eax       ; offset to current string
    mov esi, [edi]     ; extract str pointer to esi
    test esi, esi      ; test str for zero pointer
    jz _doNext         ; skip to next if empty str
    mov ecx, [esi-4]   ; extract byte length of str
    mov ebx, ecx       ; copy byte length of str
    shr ebx, 1         ; convert to char count
    mov edi, lA_len    ; copy pointer to len array
    add edi, eax       ; offset to current element
    mov [edi], ebx     ; assign str len byref
    mov edi, lA_rev    ; pointer to rev sndx array
    add edi, eax       ; offset to current element
   call sndxR          ; call sub below
    mov edi, lA_sdx    ; pointer to sndx array
    add edi, eax       ; offset to current element
   call sndx           ; call sub below
  _doNext:
    mov eax, increm
    inc eax            ; increment to next string
    mov increm, eax
    jmp _wordsLoop     ; continue

 _exitp:
    pop ebx        ; restore registers
    pop ecx        ; convention does not require ecx
    pop edx
    pop edi
    pop esi
    popf           ; restore the flags register
    mov esp, ebp   ; reset stack pointer var offset
    pop ebp        ; restore base pointer register
    ret 16         ; drop the params 16 bytes & exit
genSndx endp       ; end of procedure

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
 ;   ' Returns the 4 character Soundex code for an English word.
 ;   Public Function GetSoundexWord(sWord As String) As Long
 ;       Dim bSoundex(1 To 4) As Byte
 ;       Dim i As Long, j As Long
 ;       Dim prev As Byte
 ;       Dim code As Byte
 ;
 ;       If Len(sWord) = 0 Then Exit Function
 ;
 ;       '// Replacement
 ;       '   [a, e, h, i, o, u, w, y] = 0
 ;       '   [b, f, p, v] = 1
 ;       '   [c, g, j, k, q, s, x, z] = 2
 ;       '   [d, t] = 3
 ;       '   [l] = 4
 ;       '   [m, n] = 5
 ;       '   [r] = 6
 ;
 ;       For i = 1 To Len(sWord)
 ;           Select Case MidLcI(sWord, i) 'LCase$(Mid$(sWord, i, 1))
 ;                 ' "a", "e", "h", "i", "o", "u", "w", "y"
 ;               Case 97, 101, 104, 105, 111, 117, 119, 121:  GoTo nexti '// do nothing
 ;
 ;                 ' "b", "f", "p", "v"
 ;               Case 98, 102, 112, 118:                      code = 1 '// key labials
 ;
 ;                 ' "c", "g", "j", "k", "q", "s", "x", "z"
 ;               Case 99, 103, 106, 107, 113, 115, 120, 122:  code = 2
 ;
 ;               Case 100, 116: code = 3   ' "d", "t"
 ;               Case 108:      code = 4   ' "l"
 ;               Case 109, 110: code = 5   ' "m", "n"
 ;               Case 114:      code = 6   ' "r"
 ;           End Select
 ;
 ;           If prev <> code Then '// do nothing if most recent
 ;               j = j + 1
 ;               bSoundex(j) = code '// add new code
 ;               If j = 4 Then Exit For
 ;               prev = code
 ;           End If
 ;   nexti:
 ;       Next i
 ;
 ;       '// Return the first four values (padded with 0's)
 ;       CopyMemory GetSoundexWord, bSoundex(1), 4
 ;   End Function
;

sndx proc near   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

    mov eax, [esi-4]   ; extract byte length of str
    xor ecx, ecx       ; clear str char counter
    xor ebx, ebx       ; clear result byte offset
    xor dx, dx         ; clear soundex char codes

   ;  // Replacement
   ;     [a, e, h, i, o, u, w, y] = 0
   ;     [b, f, p, v] = 1
   ;     [c, g, j, k, q, s, x, z] = 2
   ;     [d, t] = 3
   ;     [l] = 4
   ;     [m, n] = 5
   ;     [r] = 6

  _sndxloop:
    cmp ecx, eax        ;
    je _exitsndx        ; For i = 1 To Len(sWord)

    mov dh, [esi+ecx]   ; extract current char

    cmp dh, 96          ; "a..."
    ja _lcase           ; jump to lowercase

    cmp dh, 65          ; "A"
    je _continue
    cmp dh, 66          ; "B"
    je _code_1
    cmp dh, 67          ; "C"
    je _code_2
    cmp dh, 68          ; "D"
    je _code_3
    cmp dh, 69          ; "E"
    je _continue
    cmp dh, 70          ; "F"
    je _code_1
    cmp dh, 71          ; "G"
    je _code_2
    cmp dh, 72          ; "H"
    je _continue
    cmp dh, 73          ; "I"
    je _continue
    cmp dh, 74          ; "J"
    je _code_2
    cmp dh, 75          ; "K"
    je _code_2
    cmp dh, 76          ; "L"
    je _code_4
    cmp dh, 77          ; "M"
    je _code_5
    cmp dh, 78          ; "N"
    je _code_5
    cmp dh, 79          ; "O"
    je _continue
    cmp dh, 80          ; "P"
    je _code_1
    cmp dh, 81          ; "Q"
    je _code_2
    cmp dh, 82          ; "R"
    je _code_6
    cmp dh, 83          ; "S"
    je _code_2
    cmp dh, 84          ; "T"
    je _code_3
    cmp dh, 85          ; "U"
    je _continue
    cmp dh, 86          ; "V"
    je _code_1
    cmp dh, 87          ; "W"
    je _continue
    cmp dh, 88          ; "X"
    je _code_2
    cmp dh, 89          ; "Y"
    je _continue
    cmp dh, 90          ; "Z"
    je _code_2

  _continue:
    add ecx, 2      ; increment wchar counter
    jmp _sndxloop

  _exitsndx:
    ret

  _lcase:
    cmp dh, 97          ; "a"
    je _continue
    cmp dh, 98          ; "b"
    je _code_1
    cmp dh, 99          ; "c"
    je _code_2
    cmp dh, 100         ; "d"
    je _code_3
    cmp dh, 101         ; "e"
    je _continue
    cmp dh, 102         ; "f"
    je _code_1
    cmp dh, 103         ; "g"
    je _code_2
    cmp dh, 104         ; "h"
    je _continue
    cmp dh, 105         ; "i"
    je _continue
    cmp dh, 106         ; "j"
    je _code_2
    cmp dh, 107         ; "k"
    je _code_2
    cmp dh, 108         ; "l"
    je _code_4
    cmp dh, 109         ; "m"
    je _code_5
    cmp dh, 110         ; "n"
    je _code_5
    cmp dh, 111         ; "o"
    je _continue
    cmp dh, 112         ; "p"
    je _code_1
    cmp dh, 113         ; "q"
    je _code_2
    cmp dh, 114         ; "r"
    je _code_6
    cmp dh, 115         ; "s"
    je _code_2
    cmp dh, 116         ; "t"
    je _code_3
    cmp dh, 117         ; "u"
    je _continue
    cmp dh, 118         ; "v"
    je _code_1
    cmp dh, 119         ; "w"
    je _continue
    cmp dh, 120         ; "x"
    je _code_2
    cmp dh, 121         ; "y"
    je _continue
    cmp dh, 122         ; "z"
    je _code_2
    jmp _continue

  _code_1:
    xor dl, 1
    jz _continue       ; // do nothing if most recent
    mov dl, 1          ; // add new code
    jmp _addcons

  _code_2:
    xor dl, 2
    jz _continue
    mov dl, 2
    jmp _addcons

  _code_3:
    xor dl, 3
    jz _continue
    mov dl, 3
    jmp _addcons

  _code_4:
    xor dl, 4
    jz _continue
    mov dl, 4
    jmp _addcons

  _code_5:
    xor dl, 5
    jz _continue
    mov dl, 5
    jmp _addcons

  _code_6:
    xor dl, 6
    jz _continue
    mov dl, 6

  _addcons:
    mov [edi+ebx], dl
    inc ebx
    cmp ebx, 4
    jb _continue

    ret

sndx endp           ; end of procedure

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
 ;   ' Returns the 4 character Soundex code for an English word but from right to left.
 ;   Public Function GetSoundexWordR(sWord As String) As Long
 ;       Dim bSoundex(1 To 4) As Byte
 ;       Dim i As Long, j As Long
 ;       Dim prev As Byte
 ;       Dim code As Byte
 ;
 ;       If Len(sWord) = 0 Then Exit Function
 ;
 ;       '// Replacement
 ;       '   [a, e, h, i, o, u, w, y] = 0
 ;       '   [b, f, p, v] = 1
 ;       '   [c, g, j, k, q, s, x, z] = 2
 ;       '   [d, t] = 3
 ;       '   [l] = 4
 ;       '   [m, n] = 5
 ;       '   [r] = 6
 ;
 ;       For i = Len(sWord) To 1 Step -1
 ;           Select Case MidLcI(sWord, i) 'LCase$(Mid$(sWord, i, 1))
 ;                 ' "a", "e", "h", "i", "o", "u", "w", "y"
 ;               Case 97, 101, 104, 105, 111, 117, 119, 121:  GoTo nexti '// do nothing
 ;
 ;                 ' "b", "f", "p", "v"
 ;               Case 98, 102, 112, 118:                      code = 1 '// key labials
 ;
 ;                 ' "c", "g", "j", "k", "q", "s", "x", "z"
 ;               Case 99, 103, 106, 107, 113, 115, 120, 122:  code = 2
 ;
 ;               Case 100, 116: code = 3   ' "d", "t"
 ;               Case 108:      code = 4   ' "l"
 ;               Case 109, 110: code = 5   ' "m", "n"
 ;               Case 114:      code = 6   ' "r"
 ;           End Select
 ;
 ;           If prev <> code Then '// do nothing if most recent
 ;               j = j + 1
 ;               bSoundex(j) = code '// add new code
 ;               If j = 4 Then Exit For
 ;               prev = code
 ;           End If
 ;   nexti:
 ;       Next i
 ;
 ;       '// Return the first four values (padded with 0's)
 ;       CopyMemory GetSoundexWordR, bSoundex(1), 4
 ;   End Function
;

sndxR proc near   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

    xor ebx, ebx       ; clear result byte offset
    xor dx, dx         ; clear soundex char codes

   ;  // Replacement
   ;     [a, e, h, i, o, u, w, y] = 0
   ;     [b, f, p, v] = 1
   ;     [c, g, j, k, q, s, x, z] = 2
   ;     [d, t] = 3
   ;     [l] = 4
   ;     [m, n] = 5
   ;     [r] = 6

    jmp _sndxRloop      ; jump distance fix
  _exitsndxR:
    ret

  _sndxRloop:
    jcxz _exitsndxR     ; For i = Len(sWord) To 1 Step -1
    sub ecx, 2

    mov dh, [esi+ecx]   ; extract current char

    cmp dh, 96          ; "a..."
    ja _lcaseR          ; jump to lowercase

    cmp dh, 65          ; "A"
    je _sndxRloop
    cmp dh, 66          ; "B"
    je _codeR_1
    cmp dh, 67          ; "C"
    je _codeR_2
    cmp dh, 68          ; "D"
    je _codeR_3
    cmp dh, 69          ; "E"
    je _sndxRloop
    cmp dh, 70          ; "F"
    je _codeR_1
    cmp dh, 71          ; "G"
    je _codeR_2
    cmp dh, 72          ; "H"
    je _sndxRloop
    cmp dh, 73          ; "I"
    je _sndxRloop
    cmp dh, 74          ; "J"
    je _codeR_2
    cmp dh, 75          ; "K"
    je _codeR_2
    cmp dh, 76          ; "L"
    je _codeR_4
    cmp dh, 77          ; "M"
    je _codeR_5
    cmp dh, 78          ; "N"
    je _codeR_5
    cmp dh, 79          ; "O"
    je _sndxRloop
    cmp dh, 80          ; "P"
    je _codeR_1
    cmp dh, 81          ; "Q"
    je _codeR_2
    cmp dh, 82          ; "R"
    je _codeR_6
    cmp dh, 83          ; "S"
    je _codeR_2
    cmp dh, 84          ; "T"
    je _codeR_3
    cmp dh, 85          ; "U"
    je _sndxRloop
    cmp dh, 86          ; "V"
    je _codeR_1
    cmp dh, 87          ; "W"
    je _sndxRloop
    cmp dh, 88          ; "X"
    je _codeR_2
    cmp dh, 89          ; "Y"
    je _sndxRloop
    cmp dh, 90          ; "Z"
    je _codeR_2

  _continueR:
    jmp _sndxRloop

  _lcaseR:
    cmp dh, 97          ; "a"
    je _sndxRloop
    cmp dh, 98          ; "b"
    je _codeR_1
    cmp dh, 99          ; "c"
    je _codeR_2
    cmp dh, 100         ; "d"
    je _codeR_3
    cmp dh, 101         ; "e"
    je _sndxRloop
    cmp dh, 102         ; "f"
    je _codeR_1
    cmp dh, 103         ; "g"
    je _codeR_2
    cmp dh, 104         ; "h"
    je _sndxRloop
    cmp dh, 105         ; "i"
    je _sndxRloop
    cmp dh, 106         ; "j"
    je _codeR_2
    cmp dh, 107         ; "k"
    je _codeR_2
    cmp dh, 108         ; "l"
    je _codeR_4
    cmp dh, 109         ; "m"
    je _codeR_5
    cmp dh, 110         ; "n"
    je _codeR_5
    cmp dh, 111         ; "o"
    je _sndxRloop
    cmp dh, 112         ; "p"
    je _codeR_1
    cmp dh, 113         ; "q"
    je _codeR_2
    cmp dh, 114         ; "r"
    je _codeR_6
    cmp dh, 115         ; "s"
    je _codeR_2
    cmp dh, 116         ; "t"
    je _codeR_3
    cmp dh, 117         ; "u"
    je _sndxRloop
    cmp dh, 118         ; "v"
    je _codeR_1
    cmp dh, 119         ; "w"
    je _sndxRloop
    cmp dh, 120         ; "x"
    je _codeR_2
    cmp dh, 121         ; "y"
    je _sndxRloop
    cmp dh, 122         ; "z"
    je _codeR_2

    jmp _sndxRloop

  _codeR_1:
    xor dl, 1
    jz _continueR   ; do nothing if most recent
    mov dl, 1       ; add new code
    jmp _addconsR

  _codeR_2:
    xor dl, 2
    jz _continueR
    mov dl, 2
    jmp _addconsR

  _codeR_3:
    xor dl, 3
    jz _continueR
    mov dl, 3
    jmp _addconsR

  _codeR_4:
    xor dl, 4
    jz _continueR
    mov dl, 4
    jmp _addconsR

  _codeR_5:
    xor dl, 5
    jz _continueR
    mov dl, 5
    jmp _addconsR

  _codeR_6:
    xor dl, 6
    jz _continueR
    mov dl, 6

  _addconsR:
    mov [edi+ebx], dl
    inc ebx
    cmp ebx, 4
    jb _continueR

    ret

sndxR endp          ; end of procedure

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

end genSndx         ; end of code, indicate entry point

; A big thanks to Robert for your generous help with assembler.

; Whenever I had problems trying to figure out how to get this
; to work I was able to find an example that demonstrated just
; the right solution and it was always one of your submissions.

; I thought my VB code that did this op was pretty fast... :)


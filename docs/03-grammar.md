# 03 - Grammar ANTLR

Arquivo: `grammar/rpa.g4`

```antlr
grammar rpa;

// ========== PROGRAMA ==========
programa
    : (comando NEWLINE*)* EOF
    ;

// ========== COMANDOS ==========
comando
    : blocoIf          # CmdIf
    | blocoWhile       # CmdWhile
    | blocoFor         # CmdFor
    | blocoTry         # CmdTry
    | blocoDef         # CmdDef
    | cmdGenerico      # CmdGenerico
    ;

// --- Comando genérico: qualquer ID + argumentos ---
cmdGenerico
    : ID (arg)*
    ;

arg
    : '--' ID '=' expressao          # ArgFlag
    | expressao                      # ArgPosicional
    ;

// --- If / Elseif / Else ---
blocoIf
    : 'if' expressao NEWLINE*
      (comando NEWLINE*)*
      ('elseif' expressao NEWLINE* (comando NEWLINE*)*)*
      ('else' NEWLINE* (comando NEWLINE*)*)?
      'endif'
    ;

// --- While ---
blocoWhile
    : 'while' expressao NEWLINE*
      (comando NEWLINE*)*
      'endwhile'
    ;

// --- For ---
blocoFor
    : 'for' VAR '=' expressao 'to' expressao ('step' expressao)? NEWLINE*
      (comando NEWLINE*)*
      'next'
    ;

// --- Try / Catch / Finally ---
blocoTry
    : 'try' NEWLINE*
      (comando NEWLINE*)*
      ('catch' (VAR)? NEWLINE* (comando NEWLINE*)*)?
      ('finally' NEWLINE* (comando NEWLINE*)*)?
      'endtry'
    ;

// --- Def (funções) ---
blocoDef
    : 'def' ID '(' (VAR (',' VAR)*)? ')' NEWLINE*
      (comando NEWLINE*)*
      'enddef'
    ;

// ========== EXPRESSÕES ==========
expressao
    : '(' expressao ')'
    | '!' expressao
    | expressao ('==' | '!=' | '>' | '<' | '>=' | '<=') expressao
    | expressao ('+' | '-' | 'or') expressao
    | expressao ('*' | '/' | 'and') expressao
    | NUM
    | STRING
    | BOOL
    | VAR
    | listLiteral
    | mapLiteral
    ;

listLiteral : '[' (expressao (',' expressao)*)? ']' ;
mapLiteral  : '{' (STRING ':' expressao (',' STRING ':' expressao)*)? '}' ;

// ========== TOKENS ==========
VAR     : '$' [a-zA-Z_][a-zA-Z_0-9]* ;
ID      : [a-zA-Z_][a-zA-Z_0-9]* ;
NUM     : [0-9]+ ('.' [0-9]+)? ;
STRING  : '"' (~["])* '"' ;
BOOL    : 'true' | 'false' ;
NEWLINE : [\r\n]+ ;
WS      : [ \t]+ -> skip ;
COMMENT : '#' ~[\r\n]* -> skip ;
```

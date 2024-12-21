Dim opcaoPc, opcaoJogador, resp
Dim perdeu, ganhou, empate, ponto
perdeu = 0
ganhou = 0
empate = 0
ponto = 1
call jogo
sub jogo()
randomize(second(time))
opcaoPc=int(rnd*3)+1
opcaoJogador=cint(inputbox( "PLACAR ATUAL: " + vbnewline & _
                            "VITORIAS: "& ganhou &"" + vbnewline & _
                            "DERROTAS: "& perdeu &"" + vbnewline & _
                            "EMPATES: "& empate &"" + vbnewline & _
                            "" + vbnewline & _
                            "OPCOES: " + vbnewline & _
                            "[1] PEDRA" + vbnewline & _
                            "[2] PAPEL" + vbnewline & _
                            "[3] TESOURA" + vbnewline & _
                            "[4] SAIR DO JOGO" + vbnewline & _
                            "" + vbnewline & _
                            "Digite o numero de uma opcao:", "JOKENPO"))

if opcaoJogador = 4 then
    sairJogo()
elseif opcaoJogador=1 or opcaoJogador=2 or opcaoJogador=3 then
    select case opcaoPc
        case 1:
            pedraPC()
        case 2:
            papelPC()
        case 3:
            tesouraPC()
    end select
else
    opcaoInvalida()
end if
end sub

function pedraPC
if opcaoJogador=1 then
    empate = empate + 1
    resp=msgbox("EMPATE!" + vbnewline & _
                "Voce jogou [1] PEDRA e o computador tambem!", vbinformation + vbokonly)
                pergunta()
elseif opcaoJogador=3 then
    perdeu = perdeu + 1
    resp=msgbox("VOCE PERDEU!" + vbnewline & _
                "Voce jogou [3] TESOURA e o computador PEDRA!", vbinformation + vbokonly)
                pergunta()
elseif opcaoJogador=2 then
    ganhou = ganhou + 1
    resp=msgbox("VOCE GANHOU!" + vbnewline & _
                "Voce jogou [2] PAPEL e o computador PEDRA!", vbinformation + vbokonly)
                pergunta()
else
    opcaoInvalida()
end if
end function

function papelPC
if opcaoJogador=2 then
    empate=empate + 1
    resp=msgbox("EMPATE!" + vbnewline & _
                "Voce jogou [2] PAPEL e o computador tambem!", vbinformation + vbokonly)
                pergunta()
elseif opcaoJogador=1 then
    perdeu=perdeu + 1
    resp=msgbox("VOCE PERDEU!" + vbnewline & _
                "Voce jogou [1] PEDRA e o computador PAPEL!", vbinformation + vbokonly)
                pergunta()
elseif opcaoJogador=3 then
    ganhou=ganhou + 1
    resp=msgbox("VOCE GANHOU!" + vbnewline & _
                "Voce jogou [3] TESOURA e o computador PAPEL!", vbinformation + vbokonly)
                pergunta()
else
    opcaoInvalida()
end if
end function

function tesouraPC
if opcaoJogador=3 then
    empate=empate + 1
    resp=msgbox("EMPATE!" + vbnewline & _
                "Voce jogou [3] TESOURA e o computador tambem!", vbinformation + vbokonly)
                pergunta()
elseif opcaoJogador=2 then
    perdeu=perdeu + 1
    resp=msgbox("VOCE PERDEU!" + vbnewline & _
                "Voce jogou [2] PAPEL e o computador TESOURA!", vbinformation + vbokonly)
                pergunta()
elseif opcaoJogador=1 then
    ganhou=ganhou + 1
    resp=msgbox("VOCE GANHOU!" + vbnewline & _
                "Voce jogou [1] PEDRA e o computador TESOURA!", vbinformation + vbokonly)
                pergunta()
else
    opcaoInvalida()
end if
end function

function pergunta
resp=msgbox("Deseja jogar novamente?", vbquestion + vbyesno)
        if resp=vbyes then
            call jogo
        else
            sairJogo()
        end if
end function

function opcaoInvalida
    resp=msgbox("Digite uma opcao valida!", vbinformation + vbokonly)
    pergunta()
end function

function sairJogo
    resp=msgbox("Voce realmente deseja sair do jogo?", vbquestion + vbyesno)
    if resp=vbyes then
        wscript.quit
    else
        call jogo
    end if
end function
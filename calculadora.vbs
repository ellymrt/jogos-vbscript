Dim num1, num2, sinal, resposta
Dim conta, acertou, errou

acertou=0
errou=0

call calculadora
sub calculadora()

randomize(second(time))
num1=int(rnd*100)
num2=int(rnd*100)
sinal=int(rnd*3)+1

select case sinal
    case 1:
        adicao()
    case 2:
        subtracao()
    case 3:
        multiplicacao()
end select
end sub

function adicao
    conta=num1+num2
    resposta=cInt(inputbox("QUANTO EH: "& num1 &" + "& num2 &""))
    if conta=resposta then
        acertou=acertou+1
        certo()
    else
        errou=errou+1
        erro()
    end if
end function

function subtracao
    conta=num1-num2
    resposta=cInt(inputbox("QUANTO EH: "& num1 &" - "& num2 &""))
    if resposta=conta then
        acertou=acertou+1
        certo()
    else
        errou=errou+1
        erro()
    end if
end function

function multiplicacao
    conta=(num1*num2)
    resposta=cInt(inputbox("QUANTO EH: "& num1 &" * "& num2 &""))
    if resposta=conta then
        acertou=acertou+1
        certo()
    else
        errou=errou+1
        erro()
    end if
end function

function certo
    resposta=msgbox("VOCE ACERTOU!" + vbnewline & _
                    "A resposta era: "& conta &"" + vbnewline & _
                    "E voce colocou: "& resposta &"" + vbnewline & _
                    "" + vbnewline & _
                    "ERROS: "& errou &"" + vbnewline & _
                    "ACERTOS: "& acertou &"" + vbnewline & _
                    "" + vbnewline & _
                    "Deseja jogar novamente?", vbquestion + vbyesno, "ATENCAO")
    if resposta=vbyes then
        call calculadora
    else
        wscript.quit
    end if
end function

function erro
    resposta=msgbox("VOCE ERROU!" + vbnewline & _
                    "A resposta era: "& conta &"" + vbnewline & _
                    "E voce colocou: "& resposta &"" + vbnewline & _
                    "" + vbnewline & _
                    "ERROS: "& errou &"" + vbnewline & _
                    "ACERTOS: "& acertou &"" + vbnewline & _
                    "" + vbnewline & _
                    "Deseja jogar novamente?", vbquestion + vbyesno, "ATENCAO")
    if resposta=vbyes then
        call calculadora
    else
        wscript.quit
    end if
end function
Dim palpite, tentativa, sorteado
call sorteio
sub sorteio()
randomize(second(time))
sorteado=int(rnd*50)+1
for tentativa=1 to 5 step 1
    palpite=cint(inputbox("Digite seu palpite: " + vbnewline & _
                          "Tentativas: "& tentativa &"", "ADIVINHA"))
    if palpite=sorteado then
        resp=msgbox("Parabens vc ganhou! Deseja jogar novamente?", vbyesno+vbquestion)
        if resp=vbyes then
            call sorteio
        else
            exit sub
        end if
    end if
next
msgbox("Fim do jogo!"),vbinformation+vbokonly,"AVISO"
end sub
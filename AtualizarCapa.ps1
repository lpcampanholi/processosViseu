function CopiarArquivosAntigos {
  param($NomeDoAutor, $NomeDoLivro, $Diretorio, $Formato)

  # Verifica se a pasta _Fisico existe
  if (Test-Path -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Fisico\" -PathType Container) {
    # Renomeia a pasta para _Físico
    Rename-Item -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Fisico\" -NewName "_Físico" -Force
    Write-Host "Pasta _Fisico renomeada para _Físico."
  }

  Write-Host "Copiando Arquivos do Drive..."
  Copy-Item -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico\" -Destination "C:\Users\Viseu\Desktop\$NomeDoLivro" -Recurse

  if ($Formato -eq '140x210') {
  Write-Host "Copiando Gabarito de Capa 140x210..."
  Copy-Item -Path "G:\Meu Drive\_Gabaritos Capa\Gabarito_Capa_14x21cm_2023 Folder\Gabarito_Capa_14x21cm.indd" -Destination "C:\Users\Viseu\Desktop\$NomeDoLivro"
  Write-Host "Abrindo Gabarito 140x210..."
  Invoke-Item -Path "C:\Users\Viseu\Desktop\$NomeDoLivro\Gabarito_Capa_14x21cm.indd"
}

  if ($Formato -eq '160x230') {
  Write-Host "Copiando Gabarito de Capa 160x230..."
  Copy-Item -Path "G:\Meu Drive\_Gabaritos Capa\Gabarito_Capa_16x23cm_2023 Folder\Gabarito_Capa_16x23cm_2023.indd" -Destination "C:\Users\Viseu\Desktop\$NomeDoLivro"
  Write-Host "Abrindo Gabarito 160x230..."
  Invoke-Item -Path "C:\Users\Viseu\Desktop\$NomeDoLivro\Gabarito_Capa_16x23cm_2023.indd"
}

  Write-Host "Abrindo Pasta do Livro em Desktop..."
  Invoke-Item -Path "C:\Users\Viseu\Desktop\$NomeDoLivro"

  Write-Host "Abrindo arquivos .ai..."
  Get-ChildItem -Path C:\Users\Viseu\Desktop\$NomeDoLivro -Filter "*.ai" -File | ForEach-Object {
    Invoke-Item $_.FullName
  }

}


function SalvarArquivosAtualizados {
  
  param($NomeDoAutor, $NomeDoLivro, $Diretorio, $PastaCapaNova)
  
  #Criar pasta "_Descarte" em "_Físico" se não existir
  if (-not (Test-Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico\_Descarte")) {
    New-Item -ItemType Directory -Name "_Descarte" -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico"
  }

  #Mover arquivos não indd para "_Descarte"
  Get-ChildItem -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico" | Where-Object { $_.FullName -notlike "*\_Descarte\*" -and $_.Extension -ne ".indd" } | Move-Item -Destination "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico\_Descarte"

  #Copiar pasta da capa nova para o Drive em "_Físico"
  Copy-Item -Path "C:\Users\Viseu\Desktop\$NomeDoLivro\$pastaCapaNova" -Destination "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico" -Recurse

  #Criar pasta "_Descarte" em "_Impressão" se não existir
  if (-not (Test-Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão\_Descarte")) {
    New-Item -ItemType Directory -Name "_Descarte" -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão"
  }

  #Mover capa antiga para pasta "_Descarte" em "_Impressão"
  Get-ChildItem -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão" -File | Where-Object { $_.Name -match "capa" } | Move-Item -Destination "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão\_Descarte"

  #Copia nova capa para o Drive em "_Impressão"
  Get-ChildItem -Path "C:\Users\Viseu\Desktop\$NomeDoLivro\$pastaCapaNova" -Filter "*.pdf" -Recurse | Copy-Item -Destination "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão"

  #Abre pasta "_Físico" do Drive para Conferir
  Invoke-Item -Path "G:\Drives compartilhados\$Diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico"
  
}


function SelecionarDiretorio {

  Write-Host "Selecione o caminho para o projeto:"
  Write-Host "1: Projetos finalizados I"
  Write-Host "2: Projetos finalizados II"
  Write-Host "3: Viseu Criação\Projetos"

  $opcao = Read-Host "Digite o número da opção escolhida (1-3)"
  switch ($opcao) {
    '1' { return 'Projetos finalizados I' }
    '2' { return 'Projetos finalizados II' }
    '3' { return 'Viseu Criação\Projetos' }
    default {
      Write-Host "Opção inválida. Tente novamente."
      return SelecionarDiretorio
    }
  }
}


function SelecionarFormato {

  Write-Host "Selecione o formato do projeto:"
  Write-Host "1: 140x210"
  Write-Host "2: 160x230"
  
  $opcao = Read-Host "Digite o número da opção escolhida (1-2)"
  switch ($opcao) {
    '1' { return '140x210' }
    '2' { return '160x230' }
    default {
      Write-Host "Opção inválida. Tente novamente."
      return SelecionarFormato
    }
  }
}


$Diretorio = SelecionarDiretorio
$Formato = SelecionarFormato
$NomeDoAutor = Read-Host "Nome do Autor"
$NomeDoLivro = Read-Host "Nome do Livro"

CopiarArquivosAntigos -NomeDoAutor $NomeDoAutor -NomeDoLivro $NomeDoLivro -Diretorio $Diretorio -Formato $Formato

Write-Host "ATUALIZE A CAPA"
$PastaCapaNova = Read-Host "Insira o nome da Pasta da Capa Nova"

SalvarArquivosAtualizados -NomeDoAutor $NomeDoAutor -NomeDoLivro $NomeDoLivro -Diretorio $Diretorio -PastaCapaNova $PastaCapaNova

function CopiarArquivosAntigos {
  param($NomeDoAutor, $NomeDoLivro, $Diretorio)

  Write-Host "Copiando Arquivos do Drive..."
  Copy-Item -Path "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico\" -Destination "C:\Users\Viseu\Desktop\$NomeDoLivro" -Recurse

  Write-Host "Copiando Gabarito de Capa 14x21"
  Copy-Item -Path "G:\Meu Drive\_Gabaritos Capa\Gabarito_Capa_14x21cm_2023 Folder\Gabarito_Capa_14x21cm.indd" -Destination "C:\Users\Viseu\Desktop\$NomeDoLivro"

  Write-Host "Copiando Gabarito de Capa 16x23"
  Copy-Item -Path "G:\Meu Drive\_Gabaritos Capa\Gabarito_Capa_16x23cm_2023 Folder\Gabarito_Capa_16x23cm_2023.indd" -Destination "C:\Users\Viseu\Desktop\$NomeDoLivro"

  White-Host "Abrindo Pasta do Livro em Desktop"
  Invoke-Item -Path "C:\Users\Viseu\Desktop\$NomeDoLivro"

  Write-Host "Abrindo Gabarito 14x21"
  Invoke-Item -Path "C:\Users\Viseu\Desktop\$NomeDoLivro\Gabarito_Capa_14x21cm.indd"

  Write-Host "Abrindo arquivos .ai"
  Get-ChildItem -Path C:\Users\Viseu\Desktop\$NomeDoLivro -Filter "*.ai" -File | ForEach-Object {
    Invoke-Item $_.FullName
  }

}


function SalvarArquivosAtualizados {
  
  param($NomeDoAutor, $NomeDoLivro, $Diretorio, $PastaCapaNova)
  
  #Criar pasta "_Descarte" em "_Físico"
  New-Item -ItemType Directory -Name "_Descarte" -Path "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico"

  #Mover arquivos para pasta "_Descarte" 
  Get-ChildItem -Path "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico" | Where-Object { $_.FullName -notlike "*\_Descarte\*" -and $_.Extension -ne ".indd" } | Move-Item -Destination "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico\_Descarte"

  #Copiar pasta da capa nova para o Drive em "_Físico"
  Copy-Item -Path "C:\Users\Viseu\Desktop\$NomeDoLivro\$pastaCapaNova" -Destination "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico" -Recurse

  #Criar pasta "_Descarte" em "_Impressão"
  New-Item -ItemType Directory -Name "_Descarte" -Path "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão"

  #Mover capa antiga para pasta "_Descarte" em "_Impressão"
  Get-ChildItem -Path "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão" -File | Where-Object { $_.Name -match "capa" } | Move-Item -Destination "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão\_Descarte"

  #Copia nova capa para o Drive em "_Impressão"
  Get-ChildItem -Path "C:\Users\Viseu\Desktop\$NomeDoLivro\$pastaCapaNova" -Filter "*.pdf" -Recurse | Copy-Item -Destination "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Impressão"

  #Abre pasta "_Físico" do Drive para Conferir
  Invoke-Item -Path "G:\Drives compartilhados\$diretorio\$NomeDoAutor\$NomeDoLivro\_Arte\_Físico"
  
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

$Diretorio = SelecionarDiretorio
$NomeDoAutor = Read-Host "Nome do Autor"
$NomeDoLivro = Read-Host "Nome do Livro"

CopiarArquivosAntigos -NomeDoAutor $NomeDoAutor -NomeDoLivro $NomeDoLivro -Diretorio $Diretorio

$PastaCapaNova = Read-Host "Insira o nome da Pasta da Capa Nova"

SalvarArquivosAtualizados -NomeDoAutor $NomeDoAutor -NomeDoLivro $NomeDoLivro -Diretorio $Diretorio -PastaCapaNova $PastaCapaNova

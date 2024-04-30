//////////PROCESSO ATUALIZAR CAPA



#Copiar Livro para o PC
#Atualização de Capa


$nomeDoLivro = 'Fronteiras sobre tormenta'
$nomeDoAutor = 'Manuel Torres Faria (M. T. Faria)'
$diretorio = 'Viseu Criação\Projetos'

# Diretórios: 'Projetos finalizados I', 'Projetos finalizados II', 'Viseu Criação\Projetos'

#Copiar Livro
Copy-Item -Path "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Arte\_Físico\" -Destination "C:\Users\Viseu\Desktop\$nomeDoLivro" -Recurse

#Copiar Gabarito de Capa 14x21
Copy-Item -Path "G:\Meu Drive\_Gabaritos Capa\Gabarito_Capa_14x21cm_2023 Folder\Gabarito_Capa_14x21cm.indd" -Destination "C:\Users\Viseu\Desktop\$nomeDoLivro"

#Copiar Gabarito de Capa 16x23
Copy-Item -Path "G:\Meu Drive\_Gabaritos Capa\Gabarito_Capa_16x23cm_2023 Folder\Gabarito_Capa_16x23cm_2023.indd" -Destination "C:\Users\Viseu\Desktop\$nomeDoLivro"

#Abre Pasta do Livro em "Desktop"
Invoke-Item -Path "C:\Users\Viseu\Desktop\$nomeDoLivro"

#Abre Gabarito 14x21
Invoke-Item -Path "C:\Users\Viseu\Desktop\$nomeDoLivro\Gabarito_Capa_14x21cm.indd"

#Abre os arquivos .ai
Get-ChildItem -Path C:\Users\Viseu\Desktop\$nomeDoLivro -Filter "*.ai" -File | ForEach-Object {
    Invoke-Item $_.FullName
}






[Atualizar Capa]





#Copiar Livro para o Drive

$pastaCapaNova = '9788530006648_capa'

#Criar pasta "_Descarte" em "_Físico"
New-Item -ItemType Directory -Name "_Descarte" -Path "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Arte\_Físico"

#Mover arquivos para pasta "_Descarte" 
Get-ChildItem -Path "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Arte\_Físico" | Where-Object { $_.FullName -notlike "*\_Descarte\*" -and $_.Extension -ne ".indd" } | Move-Item -Destination "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Arte\_Físico\_Descarte"

#Copiar pasta da capa nova para o Drive em "_Físico"
Copy-Item -Path "C:\Users\Viseu\Desktop\$nomeDoLivro\$pastaCapaNova" -Destination "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Arte\_Físico" -Recurse

#Criar pasta "_Descarte" em "_Impressão"
New-Item -ItemType Directory -Name "_Descarte" -Path "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Impressão"

#Mover capa antiga para pasta "_Descarte" em "_Impressão"
Get-ChildItem -Path "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Impressão" -File | Where-Object { $_.Name -match "capa" } | Move-Item -Destination "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Impressão\_Descarte"

#Copia nova capa para o Drive em "_Impressão"
Get-ChildItem -Path "C:\Users\Viseu\Desktop\$nomeDoLivro\$pastaCapaNova" -Filter "*.pdf" -Recurse | Copy-Item -Destination "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Impressão"

#Abre pasta "_Físico" do Drive para Conferir
Invoke-Item -Path "G:\Drives compartilhados\$diretorio\$nomeDoAutor\$nomeDoLivro\_Arte\_Físico"




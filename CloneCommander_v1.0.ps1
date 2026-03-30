# ==============================================================================
# --- MECANISMO ANTI-INSTÂNCIA (Ativado para Produção) ---
# ==============================================================================
$mutexName = "Global\CloneCommander"
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)

# (Mantenha o seu 'if (-not $createdNew)' que faz o script fechar logo aqui embaixo)

# ==============================================================================
# --- MOTOR DE DIRETÓRIO HÍBRIDO (.PS1 ou .EXE) ---
# ==============================================================================
$global:AppRoot = if ($PSScriptRoot) { $PSScriptRoot } else { [System.AppDomain]::CurrentDomain.BaseDirectory }

$rootDir = $global:AppRoot # Mantém a compatibilidade com o resto do código
$bookmarksDir = Join-Path $global:AppRoot "config\Bookmarks"
$recoveryDir = Join-Path $global:AppRoot "config\Tabrecovery"

# Adicionamos o "global:" para as funções acharem o arquivo de olhos fechados
$global:recoveryFile = Join-Path $rootDir "config\Tabrecovery\Recovery.json"

if (-not $createdNew) {
   Add-Type -AssemblyName System.Windows.Forms
   [System.Windows.Forms.MessageBox]::Show("O Clone Commander já está aberto.", "Aviso", 0, 48)
   exit
}

# ==============================================================================
# --- CONFIGURAÇÃO BÁSICA ---
# ==============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- 1. BACKEND DE BOOKMARKS ---
$rootDir = $global:AppRoot
$bookmarksDir = Join-Path $rootDir "config\Bookmarks"
$jsonFile = Join-Path $bookmarksDir "Bookmarks.json"

if (-not (Test-Path $bookmarksDir)) { New-Item -ItemType Directory -Path $bookmarksDir -Force | Out-Null }

if (-not (Test-Path $jsonFile)) {
    $defaultData = [PSCustomObject]@{ Name = "Root"; Type = "Root"; Children = @() }
    $jsonContent = $defaultData | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($jsonFile, $jsonContent, [System.Text.Encoding]::UTF8)
}

function Get-BookmarksData {
    try {
        $raw = [System.IO.File]::ReadAllText($jsonFile, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($raw)) { return [PSCustomObject]@{ Name="Root"; Children=@() } }
        $obj = $raw | ConvertFrom-Json
        if (-not $obj) { return [PSCustomObject]@{ Name="Root"; Children=@() } }
        return $obj
    } catch { return [PSCustomObject]@{ Name="Root"; Children=@() } }
}

function Save-BookmarksData {
    param($DataObj)
    $json = $DataObj | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($jsonFile, $json, [System.Text.Encoding]::UTF8)
}

function Test-PathIsFavorite {
    param($TargetData, $SearchPath)
    if ($TargetData.Type -eq "Link" -and $TargetData.Path -eq $SearchPath) { return $true }
    
    $hasChildren = $false
    if ($TargetData -is [System.Collections.Hashtable]) { $hasChildren = $TargetData.ContainsKey('Children') }
    else { $hasChildren = ($TargetData.PSObject.Properties.Match('Children').Count -gt 0) }

    if ($hasChildren -and $TargetData.Children) {
        foreach ($child in $TargetData.Children) {
            if (Test-PathIsFavorite -TargetData $child -SearchPath $SearchPath) { return $true }
        }
    }
    return $false
}

function Add-Favorite {
    param($Name, $Path, $ParentNodeData)
    
    $saveRoot = $false
    if ($ParentNodeData -eq $null) {
        $ParentNodeData = Get-BookmarksData
        $saveRoot = $true
    }

    $newItem = [PSCustomObject]@{ Name = $Name; Type = "Link"; Path = $Path }

    $propExists = $false
    if ($ParentNodeData -is [System.Collections.Hashtable]) { $propExists = $ParentNodeData.ContainsKey("Children") }
    else { $propExists = ($ParentNodeData.PSObject.Properties.Match('Children').Count -gt 0) }

    # Proteção rígida contra op_Addition: Sempre converte para Array antes de somar
    if (-not $propExists) {
        if ($ParentNodeData -is [PSCustomObject]) {
            $ParentNodeData | Add-Member -MemberType NoteProperty -Name "Children" -Value @($newItem)
        } else {
            $ParentNodeData["Children"] = @($newItem)
        }
    } else {
        if ($ParentNodeData.Children -eq $null) { 
            $ParentNodeData.Children = @($newItem) 
        } else { 
            $ParentNodeData.Children = @($ParentNodeData.Children) + $newItem
        }
    }

    if ($saveRoot) { Save-BookmarksData $ParentNodeData }
}

function Remove-FavoriteByPath {
    param($PathToRemove)
    $data = Get-BookmarksData
    
    function Remove-Recursive($parent) {
        if ($parent.Children) {
            $newChildren = @()
            $changed = $false
            foreach ($child in $parent.Children) {
                if ($child.Type -eq "Link" -and $child.Path -eq $PathToRemove) {
                    $changed = $true
                } else {
                    if ($child.Type -eq "Folder") { Remove-Recursive $child }
                    $newChildren += $child
                }
            }
            if ($changed -or $newChildren.Count -ne $parent.Children.Count) {
                $parent.Children = $newChildren
            }
        }
    }
    
    Remove-Recursive $data
    Save-BookmarksData $data
}

# =========================================================================
# --- SELETOR DE PASTA (AGORA COM ESTILO ÁRVORE) ---
# =========================================================================
function Show-FolderSelectionDialog {
    $selForm = New-Object System.Windows.Forms.Form
    $selForm.Text = "Selecionar Pasta"
    $selForm.Size = New-Object System.Drawing.Size(350, 450)
    $selForm.StartPosition = "CenterParent"
    $selForm.FormBorderStyle = "FixedToolWindow"

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Onde deseja salvar este marcador?"
    $lbl.Dock = "Top"; $lbl.Height = 30; $lbl.TextAlign = "MiddleCenter"
    $selForm.Controls.Add($lbl)

    # --- LISTA DE ÍCONES PARA A ÁRVORE ---
    $imgList = New-Object System.Windows.Forms.ImageList
    $imgList.ImageSize = New-Object System.Drawing.Size(16, 16)
    $imgList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
    if ($global:FolderIconBmp) { $imgList.Images.Add("Folder", $global:FolderIconBmp) }
    
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    # --- TREEVIEW SUBSTITUINDO O LISTBOX ---
    $moveTree = New-Object System.Windows.Forms.TreeView
    $moveTree.Dock = "Top"; $moveTree.Height = 320
    $moveTree.ImageList = $imgList
    $moveTree.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $moveTree.HideSelection = $false
    $selForm.Controls.Add($moveTree)

    $pnlBtns = New-Object System.Windows.Forms.Panel
    $pnlBtns.Dock = "Bottom"; $pnlBtns.Height = 60
    $selForm.Controls.Add($pnlBtns)

    $startX = 57; $btnY = 15
    $btnNew = New-Object System.Windows.Forms.Button
    $btnNew.Text = "Nova Pasta"; $btnNew.Width = 100; $btnNew.Height = 30
    $btnNew.Location = New-Object System.Drawing.Point($startX, $btnY)
    $pnlBtns.Controls.Add($btnNew)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Salvar Aqui"; $btnOk.DialogResult = "OK"; $btnOk.Width = 100; $btnOk.Height = 30
    $btnOk.Location = New-Object System.Drawing.Point(($startX + 100 + 20), $btnY)
    $btnOk.BackColor = "#0078D7"; $btnOk.ForeColor = "White"; $btnOk.FlatStyle = "Flat"
    $pnlBtns.Controls.Add($btnOk)

    $rootData = Get-BookmarksData

    function Refresh-UI {
        $moveTree.Nodes.Clear()
        $moveRoot = $moveTree.Nodes.Add("Meus Marcadores (Raiz)")
        $moveRoot.Tag = $rootData
        $moveRoot.ImageIndex = 0; $moveRoot.SelectedImageIndex = 0
        $moveRoot.NodeFont = $boldFont
        
        function Build-Tree($parentNode, $childrenData) {
            if ($null -eq $childrenData) { return }
            foreach ($item in @($childrenData)) {
                if ($item.Type -eq "Folder") {
                    $newNode = $parentNode.Nodes.Add($item.Name)
                    $newNode.Tag = $item
                    $newNode.ImageIndex = 0; $newNode.SelectedImageIndex = 0
                    $newNode.NodeFont = $boldFont
                    Build-Tree $newNode $item.Children
                }
            }
        }
        Build-Tree $moveRoot $rootData.Children
        $moveRoot.ExpandAll()
        $moveTree.SelectedNode = $moveRoot
    }
    
    Refresh-UI

    $btnNew.Add_Click({
        $selNode = $moveTree.SelectedNode
        if ($selNode -ne $null) {
            $parent = $selNode.Tag
            
            $input = New-Object System.Windows.Forms.Form
            $input.Text = "Nome da Pasta"; $input.Size = New-Object System.Drawing.Size(250, 120)
            $input.StartPosition = "CenterParent"; $input.FormBorderStyle = "FixedToolWindow"
            
            $tb = New-Object System.Windows.Forms.TextBox; $tb.Location = New-Object System.Drawing.Point(10, 10); $tb.Width=210
            $input.Controls.Add($tb)
            
            $bk = New-Object System.Windows.Forms.Button; $bk.Text="OK"; $bk.DialogResult="OK"; $bk.Location=New-Object System.Drawing.Point(140, 40)
            $input.Controls.Add($bk); $input.AcceptButton = $bk
            
            if ($input.ShowDialog() -eq "OK" -and -not [string]::IsNullOrWhiteSpace($tb.Text)) {
                $newFolder = [PSCustomObject]@{ Name = $tb.Text; Type = "Folder"; Children = @() }
                
                if ($null -eq $parent.Children) { $parent | Add-Member -NotePropertyName "Children" -NotePropertyValue @() }
                $parent.Children = @($parent.Children) + $newFolder
                
                Refresh-UI
            }
            $input.Dispose()
        }
    })

    # --- CORREÇÃO DE VAZAMENTO: Captura o resultado para poder limpar TUDO antes de sair ---
    $result = $null
    if ($selForm.ShowDialog() -eq "OK") {
        $selNode = $moveTree.SelectedNode
        if ($selNode -ne $null) {
            $result = @{ Selected = $selNode.Tag; Root = $rootData }
        }
    }
    
    $selForm.Dispose()
    $boldFont.Dispose()
    $imgList.Dispose()
    
    return $result
}

# --- FUNÇÃO SCANNER DE EXIBIÇÃO (MODERNA E BLINDADA COM LOOP) ---
function Set-View-Mode-Scanner {
    param($Browser, $ViewName)
    
    # O loop de insistência do seu código original (Tenta até 10 vezes)
    for ($i = 1; $i -le 10; $i++) {
        try {
            if ($Browser -and $Browser.ActiveXInstance -and $Browser.ActiveXInstance.Document) {
                $doc = $Browser.ActiveXInstance.Document
                
                # O Motor Moderno (IShellFolderViewDual3)
                switch ($ViewName) {
                    "Ícones Extra Grandes" { $doc.CurrentViewMode = 1; $doc.IconSize = 256 }
                    "Ícones Grandes"       { $doc.CurrentViewMode = 1; $doc.IconSize = 96 }
                    "Ícones Medios"        { $doc.CurrentViewMode = 1; $doc.IconSize = 48 }
                    "Ícones Pequenos"      { $doc.CurrentViewMode = 2; $doc.IconSize = 16 }
                    "Lista"                { $doc.CurrentViewMode = 3 }
                    "Detalhes"             { $doc.CurrentViewMode = 4 }
                    "Miniaturas"           { $doc.CurrentViewMode = 5 } 
                    "Lado a Lado"          { $doc.CurrentViewMode = 6 }
                    "Conteúdo"             { $doc.CurrentViewMode = 8 }
                }
                
                # --- FAXINA COM OBJECT ---
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                
                # Se passou pelo switch sem dar erro de carregamento, o comando funcionou!
                break
            }
        } catch {
            # O Explorer ainda está renderizando a pasta. O catch engole o erro silenciosamente.
        }
        
        # A pausa estratégica original para dar tempo ao Windows de respirar
        Start-Sleep -Milliseconds 150
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ==============================================================================
# --- GLOBAIS E INFRAESTRUTURA DE MULTI-THREADING (O ESQUADRÃO TÁTICO) ---
# ==============================================================================
$global:ActiveBrowser = $null
$global:LeftBrowserRef = $null
$global:RightBrowserRef = $null

# 1. O Cofre Blindado (Sincronizado)
# Tudo que for colocado aqui dentro pode ser lido pelas threads secundárias sem dar erro.
$global:SyncHash = [hashtable]::Synchronized(@{})

# 2. O Esquadrão Tático (Pool de Runspaces)
# Criamos 4 threads trabalhadoras prontas para a ação.
$maxThreads = 4
$global:RunspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
$global:RunspacePool.Open()


# ==============================================================================
# --- RESTAURADOR DE FOCO E SELEÇÃO (Usa a memória do próprio script) ---
# ==============================================================================
function Restore-ExplorerFocus {
    if ($global:ActiveBrowser -ne $null) {
        try {
            $global:ActiveBrowser.Focus() | Out-Null
            $shellView = $global:ActiveBrowser.ActiveXInstance.Document
            if ($shellView) {
                $folder = $shellView.Folder
                if ($folder) {
                    $self = $folder.Self
                    if ($self) {
                        $rawPath = $self.Path
                        
                        if ($global:FolderSelMemory.ContainsKey($rawPath)) {
                            $savedNames = $global:FolderSelMemory[$rawPath]
                            if ($savedNames.Count -gt 0) {
                                foreach ($name in $savedNames) {
                                    $tgt = $folder.ParseName($name)
                                    if ($tgt) { 
                                        $shellView.SelectItem($tgt, 17) 
                                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tgt) | Out-Null
                                    } 
                                }
                            }
                        }
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($self) | Out-Null
                    }
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null
            }
        } catch {}
    }
}

# ==============================================================================
# --- MOTOR DE RECOVERY: SALVAMENTO DE SESSÃO ---
# ==============================================================================
function Save-TabSession {
    # Usa a sua variável global de arquivo primeiro
    $recFile = $global:recoveryFile
    
    # Se a global falhar, usa o $PSScriptRoot (Pasta raiz exata onde o script está salvo)
    if (-not $recFile) {
        $recFile = Join-Path $global:AppRoot "config\Tabrecovery\Recovery.json"
    }
    
    $recDir = Split-Path $recFile -Parent

    if (-not (Test-Path $recDir)) {
        New-Item -ItemType Directory -Path $recDir -Force | Out-Null
    }

    [System.Collections.ArrayList]$leftPaths = @()
    [System.Collections.ArrayList]$rightPaths = @()

    # --- LADO ESQUERDO ---
    if ($global:LeftTabControl -and $global:LeftTabControl.TabPages) {
        foreach ($tab in $global:LeftTabControl.TabPages) {
            if ($tab.Controls.Count -gt 0) {
                $browser = $tab.Controls[0]
                $path = ""
                try { 
                    $shellDoc = $browser.ActiveXInstance.Document
                    if ($shellDoc -and $shellDoc.Folder) { $path = $shellDoc.Folder.Self.Path }
                } catch {}
                
                if ([string]::IsNullOrWhiteSpace($path)) { try { if ($browser.Url) { $path = $browser.Url.LocalPath } } catch {} }
                
                $view = "Detalhes"
                if ($browser.Tag -and $browser.Tag.ViewMode) { $view = $browser.Tag.ViewMode }
                
                if (-not [string]::IsNullOrWhiteSpace($path)) { 
                    $leftPaths.Add(@{ Path = $path; View = $view }) | Out-Null 
                }
            }
        }
    }

    # --- LADO DIREITO ---
    if ($global:RightTabControl -and $global:RightTabControl.TabPages) {
        foreach ($tab in $global:RightTabControl.TabPages) {
            if ($tab.Controls.Count -gt 0) {
                $browser = $tab.Controls[0]
                $path = ""
                try { 
                    $shellDoc = $browser.ActiveXInstance.Document
                    if ($shellDoc -and $shellDoc.Folder) { $path = $shellDoc.Folder.Self.Path }
                } catch {}
                
                if ([string]::IsNullOrWhiteSpace($path)) { try { if ($browser.Url) { $path = $browser.Url.LocalPath } } catch {} }
                
                $view = "Detalhes"
                if ($browser.Tag -and $browser.Tag.ViewMode) { $view = $browser.Tag.ViewMode }
                
                if (-not [string]::IsNullOrWhiteSpace($path)) { 
                    $rightPaths.Add(@{ Path = $path; View = $view }) | Out-Null 
                }
            }
        }
    }

    if ($leftPaths.Count -gt 0 -or $rightPaths.Count -gt 0) {
        $sessionData = @{
            LeftTabs = $leftPaths.ToArray()
            RightTabs = $rightPaths.ToArray()
        }
        $sessionData | ConvertTo-Json -Depth 4 | Set-Content -Path $recFile -Encoding UTF8 -Force
    }
}

# ==============================================================================
# --- MOTOR DE RECOVERY: LEITURA DE SESSÃO ---
# ==============================================================================
function Get-TabSession {
    try {
        # Tenta pegar da global
        $recFile = $global:recoveryFile
        
        # Fallback portátil para a pasta raiz do script
        if (-not $recFile) {
            $recFile = Join-Path $global:AppRoot "config\Tabrecovery\Recovery.json"
        }

        if (Test-Path $recFile) {
            $rawJson = [System.IO.File]::ReadAllText($recFile, [System.Text.Encoding]::UTF8)
            if (-not [string]::IsNullOrWhiteSpace($rawJson)) {
                $sessionData = $rawJson | ConvertFrom-Json
                return $sessionData
            }
        }
    } catch {}
    return $null
}

# --- HELPER: FORMATAR TAMANHO ---
function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

# =========================================================================
# --- EXTRATOR DE ÍCONES PARA O GERENCIADOR ---
# =========================================================================
$csharpIcons = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;

public static class BookmarkIcons {
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int ExtractIconEx(string szFileName, int nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, int nIcons);
    
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);

    public static Bitmap GetIcon(string file, int index) {
        IntPtr large, small;
        ExtractIconEx(file, index, out large, out small, 1);
        Bitmap bmp = null;
        if (small != IntPtr.Zero) {
            using (Icon ico = Icon.FromHandle(small)) { bmp = ico.ToBitmap(); }
            DestroyIcon(small);
        }
        if (large != IntPtr.Zero) DestroyIcon(large);
        return bmp;
    }
}
'@

if (-not ([System.Management.Automation.PSTypeName]'BookmarkIcons').Type) {
    try { Add-Type -TypeDefinition $csharpIcons -ReferencedAssemblies "System.Drawing" -Language CSharp -ErrorAction SilentlyContinue } catch {}
}

# Carrega a Pasta (3) e o Link (297) na memória global
if ($null -eq $global:FolderIconBmp) { try { $global:FolderIconBmp = [BookmarkIcons]::GetIcon("shell32.dll", 3) } catch {} }
if ($null -eq $global:LinkIconBmp) { try { $global:LinkIconBmp = [BookmarkIcons]::GetIcon("shell32.dll", 297) } catch {} }

# Carrega a Estrela (43) e aplica o Contorno (Outline) Escuro
if ($null -eq $global:StarIconBmp) { 
    try { 
        $origBmp = [BookmarkIcons]::GetIcon("shell32.dll", 43) 
        
        # Função para desenhar o contorno preto em volta da imagem transparente
        function Add-Outline($bmpOriginal, $isGray) {
            $w = $bmpOriginal.Width; $h = $bmpOriginal.Height
            $newBmp = New-Object System.Drawing.Bitmap($w, $h)
            $g = [System.Drawing.Graphics]::FromImage($newBmp)
            
            # Cria a silhueta preta perfeita
            $cmBlack = New-Object System.Drawing.Imaging.ColorMatrix
            $cmBlack.Matrix00 = 0; $cmBlack.Matrix11 = 0; $cmBlack.Matrix22 = 0; 
            $iaBlack = New-Object System.Drawing.Imaging.ImageAttributes
            $iaBlack.SetColorMatrix($cmBlack)
            
            # Carimba a silhueta em 8 direções para criar a borda
            $coords = @(-1,0, 1,0, 0,-1, 0,1, -1,-1, 1,1, -1,1, 1,-1)
            for ($i=0; $i -lt $coords.Length; $i+=2) {
                $rect = New-Object System.Drawing.Rectangle($coords[$i], $coords[$i+1], $w, $h)
                $g.DrawImage($bmpOriginal, $rect, 0, 0, $w, $h, [System.Drawing.GraphicsUnit]::Pixel, $iaBlack)
            }
            
            # Cola a imagem original por cima da sombra
            if ($isGray) {
                $cmGray = New-Object System.Drawing.Imaging.ColorMatrix
                $cmGray.Matrix00 = 0.3; $cmGray.Matrix01 = 0.3; $cmGray.Matrix02 = 0.3
                $cmGray.Matrix10 = 0.59; $cmGray.Matrix11 = 0.59; $cmGray.Matrix12 = 0.59
                $cmGray.Matrix20 = 0.11; $cmGray.Matrix21 = 0.11; $cmGray.Matrix22 = 0.11
                $iaGray = New-Object System.Drawing.Imaging.ImageAttributes
                $iaGray.SetColorMatrix($cmGray)
                $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
                $g.DrawImage($bmpOriginal, $rect, 0, 0, $w, $h, [System.Drawing.GraphicsUnit]::Pixel, $iaGray)
                $iaGray.Dispose()
            } else {
                $g.DrawImage($bmpOriginal, 0, 0)
            }
            
            $iaBlack.Dispose()
            $g.Dispose()
            return $newBmp
        }
        
        $global:StarIconBmp = Add-Outline $origBmp $false
        $global:StarEmptyIconBmp = Add-Outline $origBmp $true
        $origBmp.Dispose()
    } catch {} 
}

# =========================================================================
# --- GERENCIADOR DE MARCADORES (ESTILO EXPLORER COM JANELA DE EDIÇÃO) ---
# =========================================================================
function Show-BookmarkManager {
    $mgrForm = New-Object System.Windows.Forms.Form
    $mgrForm.Text = "Gerenciador de Marcadores"
    $mgrForm.Size = New-Object System.Drawing.Size(650, 500)
    $mgrForm.StartPosition = "CenterParent"
    $mgrForm.FormBorderStyle = "FixedDialog"
    $mgrForm.MaximizeBox = $false

    # --- VARIÁVEIS DE NAVEGAÇÃO ---
    $global:RootData = Get-BookmarksData
    if (-not $global:RootData.Children) { $global:RootData | Add-Member -NotePropertyName "Children" -NotePropertyValue @() }
    $global:CurrentFolder = $global:RootData
    $global:PathStack = New-Object System.Collections.Generic.Stack[PSCustomObject]

    # --- LISTA DE ÍCONES ---
    $imgList = New-Object System.Windows.Forms.ImageList
    $imgList.ImageSize = New-Object System.Drawing.Size(16, 16)
    $imgList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
    
    if ($global:LinkIconBmp) { 
        $imgList.Images.Add("Link", $global:LinkIconBmp) 
    } else {
        $emptyBmp = New-Object System.Drawing.Bitmap(16, 16)
        $imgList.Images.Add("Link", $emptyBmp)
    }

    if ($global:FolderIconBmp) { 
        $imgList.Images.Add("Folder", $global:FolderIconBmp) 
    } else { 
        $emptyBmp = New-Object System.Drawing.Bitmap(16, 16)
        $imgList.Images.Add("Folder", $emptyBmp) 
    }

    $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $regularFont = New-Object System.Drawing.Font("Segoe UI", 9)

    # --- BARRA DE ENDEREÇO (TOPO) ---
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = "Top"; $topPanel.Height = 40; $topPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
    $mgrForm.Controls.Add($topPanel)

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = "Voltar"
    $btnUp.Location = New-Object System.Drawing.Point(5, 5); $btnUp.Size = New-Object System.Drawing.Size(70, 30)
    $btnUp.Enabled = $false
    $topPanel.Controls.Add($btnUp)

    $picIcon = New-Object System.Windows.Forms.PictureBox
    $picIcon.Location = New-Object System.Drawing.Point(85, 12); $picIcon.Size = New-Object System.Drawing.Size(16, 16)
    $picIcon.SizeMode = "StretchImage"
    if ($global:FolderIconBmp) { $picIcon.Image = $global:FolderIconBmp }
    $topPanel.Controls.Add($picIcon)

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(105, 10); $txtPath.Width = 360
    $txtPath.ReadOnly = $true; $txtPath.BackColor = [System.Drawing.Color]::White
    $txtPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $topPanel.Controls.Add($txtPath)

    # =========================================================================
    # --- ÁREA PRINCIPAL (LISTVIEW - COM MODO DETALHES INVISÍVEL) ---
    # =========================================================================
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 50)
    $listView.Size = New-Object System.Drawing.Size(460, 400)
    
    # Fundamental para a linha laranja funcionar:
    $listView.View = "Details"
    $listView.HeaderStyle = "None"
    $listView.Columns.Add("Nome", 430) | Out-Null
    $listView.FullRowSelect = $true
    
    $listView.SmallImageList = $imgList
    $listView.HideSelection = $false
    $listView.AllowDrop = $true
    $listView.LabelEdit = $false 
    $listView.Font = $regularFont
    $mgrForm.Controls.Add($listView)

    # --- PAINEL DE BOTÕES ---
    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Location = New-Object System.Drawing.Point(480, 50)
    $btnPanel.Size = New-Object System.Drawing.Size(140, 400)
    $mgrForm.Controls.Add($btnPanel)

    function New-MgrBtn($Txt, $Act) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $Txt; $b.Width = 130; $b.Height = 35; $b.FlatStyle = "Standard"
        $b.Add_Click($Act); $btnPanel.Controls.Add($b)
        return $b
    }

    # --- MOTOR DE RENDERIZAÇÃO ---
    function Update-View {
        $listView.Items.Clear()
        
        $pathString = "Raiz"
        foreach ($p in $global:PathStack.ToArray()) { $pathString = $p.Name + " > " + $pathString }
        if ($global:PathStack.Count -gt 0) { $pathString = $pathString.Replace("> Raiz", "> " + $global:CurrentFolder.Name) }
        $txtPath.Text = $pathString

        $btnUp.Enabled = ($global:PathStack.Count -gt 0)

        if ($global:CurrentFolder.Children) {
            $pastas = @($global:CurrentFolder.Children | Where-Object { $_.Type -eq 'Folder' })
            $links  = @($global:CurrentFolder.Children | Where-Object { $_.Type -ne 'Folder' })
            $global:CurrentFolder.Children = $pastas + $links

            foreach ($item in @($global:CurrentFolder.Children)) {
                $lvItem = New-Object System.Windows.Forms.ListViewItem($item.Name)
                $lvItem.Tag = $item
                if ($item.Type -eq 'Folder') { $lvItem.ImageIndex = 1 } else { $lvItem.ImageIndex = 0 }
                [void]$listView.Items.Add($lvItem)
            }
        }
    }

    # --- NAVEGAÇÃO BÁSICA ---
    $btnUp.Add_Click({
        if ($global:PathStack.Count -gt 0) {
            $global:CurrentFolder = $global:PathStack.Pop()
            Update-View
        }
    })

    $listView.Add_DoubleClick({
        if ($listView.SelectedItems.Count -gt 0) {
            $selItem = $listView.SelectedItems[0].Tag
            if ($selItem.Type -eq 'Folder') {
                if ($null -eq $selItem.Children) { $selItem | Add-Member -NotePropertyName "Children" -NotePropertyValue @() }
                $global:PathStack.Push($global:CurrentFolder)
                $global:CurrentFolder = $selItem
                Update-View
            }
        }
    })

    # =========================================================================
    # --- MOTOR DE ARRASTAR E SOLTAR BLINDADO (COM LINHA LARANJA E FANTASMA) ---
    # =========================================================================
    
    # 1. Cria a nossa linha laranja física
    $lineMarker = New-Object System.Windows.Forms.Panel
    $lineMarker.BackColor = [System.Drawing.Color]::DarkOrange
    $lineMarker.Height = 3 
    $lineMarker.Visible = $false
    $lineMarker.Enabled = $false 
    $listView.Controls.Add($lineMarker)

    # 2. Cria o Fantasma (Ghost) visual
    $ghostLabel = New-Object System.Windows.Forms.Label
    $ghostLabel.AutoSize = $false
    $ghostLabel.Height = 24
    $ghostLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#EBF4FF") # Azul claro suave
    $ghostLabel.ForeColor = [System.Drawing.Color]::Black
    $ghostLabel.BorderStyle = "FixedSingle"
    $ghostLabel.TextAlign = "MiddleLeft"
    $ghostLabel.Font = $regularFont
    $ghostLabel.Visible = $false
    $ghostLabel.Enabled = $false # IMPEDE QUE O FANTASMA ROUBE O MOUSE DA LISTA
    $mgrForm.Controls.Add($ghostLabel)

    $listView.Add_ItemDrag({ param($s, $e) 
        # Copia o nome do item e mostra o fantasma
        $ghostLabel.Text = "  " + $e.Item.Text
        $ghostLabel.Width = if ($e.Item.Bounds.Width -gt 150) { $e.Item.Bounds.Width } else { 150 }
        $ghostLabel.Visible = $true
        $ghostLabel.BringToFront()
        
        $s.DoDragDrop($e.Item, [System.Windows.Forms.DragDropEffects]::Move) 
        
        # Esconde o fantasma quando você soltar o botão do mouse
        $ghostLabel.Visible = $false
    })
    
    $listView.Add_DragOver({ param($s, $e) 
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
        $pt = $s.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
        
        # --- ATUALIZA A POSIÇÃO DO FANTASMA ---
        $formPt = $mgrForm.PointToClient([System.Windows.Forms.Cursor]::Position)
        # Deslocamos 15 pixels para baixo e para o lado, igual aos atalhos do Windows
        $ghostLabel.Location = New-Object System.Drawing.Point(($formPt.X + 15), ($formPt.Y + 15))
        if (-not $ghostLabel.Visible) { $ghostLabel.Visible = $true; $ghostLabel.BringToFront() }

        # --- LÓGICA DA LINHA LARANJA ---
        $targetItem = $s.GetItemAt(10, $pt.Y)
        
        if ($targetItem -ne $null) {
            $appearsAfter = ($pt.Y -gt ($targetItem.Bounds.Top + ($targetItem.Bounds.Height / 2)))
            
            $s.Tag = @{ Index = $targetItem.Index; After = $appearsAfter }
            
            [int]$markerY = 0
            if ($appearsAfter) { 
                $markerY = [int]$targetItem.Bounds.Bottom - 1 
            } else { 
                $markerY = [int]$targetItem.Bounds.Top - 1 
            }
            
            $lineMarker.Location = New-Object System.Drawing.Point(2, $markerY)
            $lineMarker.Width = $listView.Width - 10
            if (-not $lineMarker.Visible) { $lineMarker.Visible = $true }
            $lineMarker.BringToFront()
        } else {
            $lastIdx = $s.Items.Count - 1
            if ($lastIdx -ge 0) {
                $lastItem = $s.Items[$lastIdx]
                if ($pt.Y -gt $lastItem.Bounds.Bottom) {
                    $s.Tag = @{ Index = $lastIdx; After = $true }
                    [int]$markerY = [int]$lastItem.Bounds.Bottom - 1
                    
                    $lineMarker.Location = New-Object System.Drawing.Point(2, $markerY)
                    $lineMarker.Width = $listView.Width - 10
                    if (-not $lineMarker.Visible) { $lineMarker.Visible = $true }
                    $lineMarker.BringToFront()
                    return
                }
            }
            $s.Tag = $null
            $lineMarker.Visible = $false
        }
    })

    $listView.Add_DragLeave({ param($s, $e)
        $lineMarker.Visible = $false 
        $ghostLabel.Visible = $false # Esconde o fantasma se o mouse sair da janela
    })

    $listView.Add_DragDrop({ param($s, $e)
        $lineMarker.Visible = $false 
        $ghostLabel.Visible = $false # Morte instantânea ao fantasma ao soltar!
        
        $draggedUIItem = $e.Data.GetData([System.Windows.Forms.ListViewItem])
        if ($draggedUIItem -eq $null -or $s.Tag -eq $null) { return }
        
        $itemData = $draggedUIItem.Tag
        $targetIndex = $s.Tag.Index
        $appearsAfter = $s.Tag.After
        
        $s.Tag = $null 
        
        if ($targetIndex -ne -1 -and $s.Items[$targetIndex] -ne $draggedUIItem) {
            $targetData = $s.Items[$targetIndex].Tag
            
            $lista = New-Object System.Collections.ArrayList
            [void]$lista.AddRange($global:CurrentFolder.Children)
            
            $lista.Remove($itemData)
            $newIdx = $lista.IndexOf($targetData)
            
            if ($newIdx -ne -1) {
                if ($appearsAfter) { $newIdx++ }
                
                if ($newIdx -ge $lista.Count) {
                    [void]$lista.Add($itemData)
                } else {
                    $lista.Insert($newIdx, $itemData)
                }
                
                $global:CurrentFolder.Children = $lista.ToArray()
                Update-View
            }
        }
    })

    # --- JANELA DE ENTRADA DE TEXTO (NOVA PASTA / RENOMEAR) ---
    function Show-NameDialog($tituloJanela, $textoPadrao) {
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = $tituloJanela
        $dialog.Size = New-Object System.Drawing.Size(350, 150)
        $dialog.StartPosition = "CenterParent"
        $dialog.FormBorderStyle = "FixedDialog"
        $dialog.MaximizeBox = $false
        $dialog.MinimizeBox = $false

        $lblNome = New-Object System.Windows.Forms.Label
        $lblNome.Text = "Nome:"
        $lblNome.Location = New-Object System.Drawing.Point(15, 15)
        $lblNome.AutoSize = $true
        $dialog.Controls.Add($lblNome)

        $txtCaixa = New-Object System.Windows.Forms.TextBox
        $txtCaixa.Location = New-Object System.Drawing.Point(15, 35)
        $txtCaixa.Size = New-Object System.Drawing.Size(305, 25)
        $txtCaixa.Text = $textoPadrao
        $dialog.Controls.Add($txtCaixa)

        $btnConfirma = New-Object System.Windows.Forms.Button
        $btnConfirma.Text = "OK"
        $btnConfirma.DialogResult = "OK"
        $btnConfirma.Location = New-Object System.Drawing.Point(135, 75)
        $dialog.Controls.Add($btnConfirma)

        $btnCancela = New-Object System.Windows.Forms.Button
        $btnCancela.Text = "Cancelar"
        $btnCancela.DialogResult = "Cancel"
        $btnCancela.Location = New-Object System.Drawing.Point(220, 75)
        $dialog.Controls.Add($btnCancela)

        $dialog.AcceptButton = $btnConfirma
        $dialog.CancelButton = $btnCancela

        $dialog.Select()
        $txtCaixa.Select()

        $resultado = $dialog.ShowDialog()
        $textoFinal = $txtCaixa.Text
        $dialog.Dispose()

        if ($resultado -eq "OK" -and -not [string]::IsNullOrWhiteSpace($textoFinal)) {
            return $textoFinal.Trim()
        }
        return $null
    }

    # --- BOTÕES DE AÇÃO ---
    New-MgrBtn "Nova Pasta" {
        $novoNome = Show-NameDialog "Criar Nova Pasta" "Nova Pasta"
        
        if ($novoNome) {
            $novaPasta = [PSCustomObject]@{ Name=$novoNome; Type="Folder"; Children=@() }
            $global:CurrentFolder.Children = @($global:CurrentFolder.Children) + $novaPasta
            Update-View
        }
    }

    New-MgrBtn "Renomear" { 
        if ($listView.SelectedItems.Count -gt 0) { 
            $itemSelecionado = $listView.SelectedItems[0].Tag
            $nomeAtual = $itemSelecionado.Name
            
            $novoNome = Show-NameDialog "Renomear Item" $nomeAtual
            
            if ($novoNome -and $novoNome -ne $nomeAtual) {
                $itemSelecionado.Name = $novoNome
                Update-View
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Selecione um item para renomear.")
        }
    }

    New-MgrBtn "Excluir" { 
        if ($listView.SelectedItems.Count -gt 0) {
            $itemToRemove = $listView.SelectedItems[0].Tag
            $msg = ""
            
            if ($itemToRemove.Type -eq 'Folder') {
                if ($itemToRemove.Children -and $itemToRemove.Children.Count -gt 0) {
                    $msg = "Tem certeza que deseja excluir esta pasta? Ela contem marcadores e subpastas dentro!"
                } else {
                    $msg = "Tem certeza que deseja excluir esta pasta vazia?"
                }
            } else {
                $msg = "Tem certeza que deseja excluir este marcador?"
            }

            $resposta = [System.Windows.Forms.MessageBox]::Show($msg, "Confirmar Exclusão", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($resposta -eq "Yes") {
                $global:CurrentFolder.Children = @($global:CurrentFolder.Children | Where-Object { $_ -ne $itemToRemove })
                Update-View
            }
        }
    }

    New-MgrBtn "Mover para..." {
        if ($listView.SelectedItems.Count -eq 0) { 
            [System.Windows.Forms.MessageBox]::Show("Selecione um item valido para mover.")
            return 
        }
        $itemToMove = $listView.SelectedItems[0].Tag

        $moveForm = New-Object System.Windows.Forms.Form
        $moveForm.Text = "Mover: $($itemToMove.Name)"
        $moveForm.Size = New-Object System.Drawing.Size(350, 450)
        $moveForm.StartPosition = "CenterParent"
        $moveForm.FormBorderStyle = "FixedDialog"
        $moveForm.MaximizeBox = $false

        $moveTree = New-Object System.Windows.Forms.TreeView
        $moveTree.Location = New-Object System.Drawing.Point(10, 10)
        $moveTree.Size = New-Object System.Drawing.Size(315, 350)
        $moveTree.ImageList = $imgList
        $moveTree.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $moveTree.HideSelection = $false
        $moveForm.Controls.Add($moveTree)

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = "OK"
        $btnOk.DialogResult = "OK"
        $btnOk.Location = New-Object System.Drawing.Point(125, 370)
        $btnOk.Size = New-Object System.Drawing.Size(100, 30)
        $moveForm.Controls.Add($btnOk)

        $moveRoot = $moveTree.Nodes.Add("Raiz (Meus Marcadores)")
        $moveRoot.Tag = $global:RootData
        $moveRoot.ImageIndex = 1
        $moveRoot.SelectedImageIndex = 1
        $moveRoot.NodeFont = $boldFont

        function Build-MoveTree($parentNode, $childrenData) {
            if ($null -eq $childrenData) { return }
            foreach ($item in @($childrenData)) {
                if ($item.Type -eq "Folder" -and $item -ne $itemToMove) {
                    $newNode = $parentNode.Nodes.Add($item.Name)
                    $newNode.Tag = $item
                    $newNode.ImageIndex = 1
                    $newNode.SelectedImageIndex = 1
                    $newNode.NodeFont = $boldFont
                    Build-MoveTree $newNode $item.Children
                }
            }
        }
        
        Build-MoveTree $moveRoot $global:RootData.Children
        $moveRoot.ExpandAll()
        $moveTree.SelectedNode = $moveRoot

        if ($moveForm.ShowDialog() -eq "OK") {
            $selNode = $moveTree.SelectedNode
            if ($selNode -ne $null) {
                $targetFolder = $selNode.Tag
                
                $global:CurrentFolder.Children = @($global:CurrentFolder.Children | Where-Object { $_ -ne $itemToMove })
                if ($null -eq $targetFolder.Children) { $targetFolder | Add-Member -NotePropertyName "Children" -NotePropertyValue @() }
                $targetFolder.Children = @($targetFolder.Children) + $itemToMove
                Update-View
            }
        }
        $moveForm.Dispose()
    }

    New-MgrBtn "SALVAR E SAIR" {
        Save-BookmarksData $global:RootData
        [System.Windows.Forms.MessageBox]::Show("Marcadores salvos!")
        $mgrForm.Close()
    }

    Update-View
    $mgrForm.ShowDialog() | Out-Null

    # ====================================================================
    # CORREÇÃO DE VAZAMENTO: Destruição total da janela após o uso
    # ====================================================================
    $mgrForm.Dispose()
    $imgList.Dispose()
    $boldFont.Dispose()
    $regularFont.Dispose()
}

# =========================================================================
# --- MENU DE CONTEXTO (OTIMIZADO CONTRA VAZAMENTO DE FONTES) ---
# =========================================================================
# Cria as fontes globalmente UMA ÚNICA VEZ para não entupir a memória
$global:MenuFontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:MenuFontReg = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

function Build-ContextMenu($menuItems, $dataList) {
    foreach ($item in $dataList) {
        if ($item.Type -eq 'Folder') {
            
            # Aqui removemos o "[Pasta] " e deixamos apenas o nome limpo
            $subItem = $menuItems.Add($item.Name)
            $subItem.Font = $global:MenuFontBold # Usa a referência global
            
            # Aplica a imagem da pasta amarela extraída do Windows
            if ($global:FolderIconBmp) {
                $subItem.Image = $global:FolderIconBmp
            }

            if ($item.Children) {
                Build-ContextMenu $subItem.DropDownItems $item.Children
            }
        } else {
            $linkItem = $menuItems.Add($item.Name)
            $linkItem.Font = $global:MenuFontReg # Usa a referência global
            
            $targetPath = $item.Path
            $linkItem.Add_Click({ 
                
                # ====================================================================
                # SOLUÇÃO 1 (O GUARDA-COSTAS): Proteção contra caminhos indisponíveis
                # ====================================================================
                $isVirt = ($targetPath -match "^(shell|::|search\-ms)")
                
                # Ignora o teste se for uma pasta virtual (Meu Computador, Lixeira)
                if (-not $isVirt) {
                    
                    # Testa a sobrevivência do caminho exato neste milissegundo
                    if (-not (Test-Path -LiteralPath $targetPath -ErrorAction SilentlyContinue)) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "O local selecionado se encontra indisponível no momento ou deixou de existir.`n`nCaminho: $targetPath", 
                            "Marcador Indisponível", 
                            [System.Windows.Forms.MessageBoxButtons]::OK, 
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        )
                        return # Bloqueia a ordem de navegação e protege a aba!
                    }
                }

                # Se chegou aqui, o caminho existe (ou é virtual)! Navega com segurança.
                if ($global:ActiveBrowser) { $global:ActiveBrowser.Navigate($targetPath) }
                
            }.GetNewClosure())
        }
    }
}

# --- FORM PRINCIPAL (JANELA MODERNA SEM BORDAS) ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Clone Commander v1.0"
$form.Size = New-Object System.Drawing.Size(1100, 600)
$form.StartPosition = "CenterScreen"
$form.KeyPreview = $true 

# ====================================================================
# --- GERADOR DE ÍCONE DINÂMICO NATIVO (Segoe MDL2 Assets: EC50) ---
# ====================================================================
try {
    # Cria uma tela de 32x32 pixels transparente na memória
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $gfx.Clear([System.Drawing.Color]::Transparent)
    
    # --- DETECTA O TEMA DO WINDOWS (CLARO OU ESCURO) ---
    $isDarkTaskbar = $false # Por padrão, assume que é claro
    try {
        # Lê o cérebro do Windows para saber a cor da Barra de Tarefas
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $sysTheme = (Get-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -ErrorAction Stop).SystemUsesLightTheme
        if ($sysTheme -eq 0) { $isDarkTaskbar = $true } # 0 significa Modo Escuro
    } catch { } # Se der erro (Windows antigo), mantém o padrão

    # Escolhe a tinta do pincel baseada no tema
    if ($isDarkTaskbar) {
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White) # Barra Escura = Ícone Branco
    } else {
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black) # Barra Clara = Ícone Preto
    }
    
    $font = New-Object System.Drawing.Font("Segoe MDL2 Assets", 20, [System.Drawing.FontStyle]::Regular)
    
    # Centraliza o desenho na tela
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect = New-Object System.Drawing.RectangleF(0, 0, 32, 32)
    
    # Desenha o caractere EC50 (File Explorer)
    $gfx.DrawString([char]0xEC50, $font, $brush, $rect, $format)
    
    # Converte a imagem gerada para o formato Icon e aplica na Janela (Barra de Tarefas)
    $hIcon = $bmp.GetHicon()
    $form.Icon = [System.Drawing.Icon]::FromHandle($hIcon)
    
    # Limpa as ferramentas de pintura da memória
    $gfx.Dispose(); $font.Dispose(); $brush.Dispose()
} catch {}
# ====================================================================

# 1. ARRANCAMOS A BARRA ORIGINAL E LIGAMOS A BLINDAGEM DE RESIZE
$form.FormBorderStyle = "None"
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D7") # Borda azul bem fina para destacar a janela no fundo
$form.Padding = New-Object System.Windows.Forms.Padding(1, 1, 1, 1)

# === NOVO PROTETOR CONTRA O WINDOWS SNAP E BARRA DE TAREFAS ===
$form.Add_LocationChanged({
    try { $form.MaximumSize = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea.Size } catch {}
}.GetNewClosure())

# Motor C# para Arrastar, Redimensionar e Restaurar o Minimizar da Barra de Tarefas
$dragCode = @"
using System;
using System.Runtime.InteropServices;
public class WinDrag {
    public const int WM_NCLBUTTONDOWN = 0xA1;
    public const int GWL_STYLE = -16;
    public const int WS_MINIMIZEBOX = 0x00020000;

    [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
    [DllImport("user32.dll")] public static extern bool ReleaseCapture();
    
    // Ferramentas de cirurgia na memória da janela
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    
    public static void FixTaskbarMinimize(IntPtr handle) {
        int style = GetWindowLong(handle, GWL_STYLE);
        SetWindowLong(handle, GWL_STYLE, style | WS_MINIMIZEBOX);
    }
}
"@
if (-not ([System.Management.Automation.PSTypeName]'WinDrag').Type) { Add-Type -TypeDefinition $dragCode }

# === A MÁGICA: Devolve a habilidade de minimizar pelo ícone no exato momento em que a janela nasce ===
$form.Add_Load({
    [WinDrag]::FixTaskbarMinimize($form.Handle)
}.GetNewClosure())
# =====================================================================================================

# ====================================================================
# --- O COFRE BLINDADO (SYNCHASH) RECEBE A JANELA ---
# ====================================================================
$global:SyncHash.Form = $form

# LIGA O BUFFER DUPLO VIA REFLECTION
$bindingFlags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
$form.GetType().GetProperty("DoubleBuffered", $bindingFlags).SetValue($form, $true, $null)

$form.SuspendLayout()

# 2. A NOSSA BARRA SUPERIOR CUSTOMIZADA
$titleBar = New-Object System.Windows.Forms.Panel
$titleBar.Dock = "Top"
$titleBar.Height = 26  # <--- BARRA MAIS FINA
$titleBar.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3") # <--- COR CLARA (PADRÃO EXPLORER)
$form.Controls.Add($titleBar)

# Título da Janela (Área de clique para arrastar)
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Dock = "Fill"
$lblTitle.Text = "   Clone Commander v1.0"
$lblTitle.ForeColor = "Black"  # <--- TEXTO ESCURO PARA DAR CONTRASTE
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblTitle.TextAlign = "MiddleLeft"
$titleBar.Controls.Add($lblTitle)
$lblTitle.BringToFront()

# Evento Mágico de Arrasto: Diz ao Windows que o nosso título é a barra original
$DragAction = {
    if ($_.Button -eq 'Left') {
        [WinDrag]::ReleaseCapture()
        [WinDrag]::SendMessage($form.Handle, [WinDrag]::WM_NCLBUTTONDOWN, 2, 0) # 2 = Barra de Título
    }
}
$titleBar.Add_MouseDown($DragAction)
$lblTitle.Add_MouseDown($DragAction)

$lblTitle.Add_DoubleClick({
    if ($form.WindowState -eq 'Maximized') { 
        $form.WindowState = 'Normal' 
    } else { 
        # Calcula a área útil do monitor atual pelo Tamanho Máximo
        $form.MaximumSize = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea.Size
        $form.WindowState = 'Maximized' 
    }
})

# Agrupador dos botões de Janela na Direita
$pnlWindowBtns = New-Object System.Windows.Forms.Panel
$pnlWindowBtns.Dock = "Right"
$pnlWindowBtns.Width = 180
$pnlWindowBtns.BackColor = [System.Drawing.Color]::Transparent
$titleBar.Controls.Add($pnlWindowBtns)

# A MÁGICA: Agora a função recebe a posição X exata e o tamanho manual
function New-TitleBtn($Text, $ColorHover, $FontSize, $PosX) {
    $b = New-Object System.Windows.Forms.Button
    $b.Location = New-Object System.Drawing.Point($PosX, 0) # Posição cravada
    $b.Size = New-Object System.Drawing.Size(45, 26)        # Tamanho cravado
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.ForeColor = "Black" 
    $b.Font = New-Object System.Drawing.Font("Segoe MDL2 Assets", $FontSize)
    $b.Text = $Text
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    $b.Add_MouseEnter({ $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml($ColorHover) }.GetNewClosure())
    $b.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::Transparent })
    return $b
}

# 1. BOTÃO DE AJUDA (?) - Posição X: 0
$btnHelpApp = New-TitleBtn "?" "#E5E5E5" 12 0
$btnHelpApp.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$pnlWindowBtns.Controls.Add($btnHelpApp)

$btnHelpApp.Add_Click({
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = "Guia de Atalhos"
    $helpForm.Size = New-Object System.Drawing.Size(650, 480)
    $helpForm.StartPosition = "CenterParent"
    $helpForm.FormBorderStyle = "FixedDialog"
    $helpForm.MaximizeBox = $false
    $helpForm.MinimizeBox = $false
    $helpForm.BackColor = "#F9F9F9"
    
    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Dock = "Fill"
    $rtb.ReadOnly = $true
    $rtb.BackColor = "#F9F9F9"
    $rtb.BorderStyle = "None"
    
    # Textos
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $rtb.AppendText("Guia de Atalhos e Funcionalidades Exclusivas`n`n")
    
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $rtb.AppendText("Manipulação de Abas`n")
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $rtb.AppendText("• Ctrl + T: Duplica a aba atual no mesmo painel.`n")
    $rtb.AppendText("• Ctrl + Shift + T: Abre a pasta atual no painel oposto.`n")
    $rtb.AppendText("• Botões Laterais do Mouse: Voltam ou avançam a navegação da pasta atual.`n`n")
    
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $rtb.AppendText("Renomeação Sequencial`n")
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $rtb.AppendText("• Durante a renomeação de um arquivo, não pressione Enter. Utilize a Seta para Baixo para salvar e iniciar automaticamente a edição do arquivo seguinte. Utilize a Seta para Cima para salvar e editar o arquivo anterior.`n`n")
    
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $rtb.AppendText("Modo Turbo (Painel Central)`n")
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $rtb.AppendText("• Ao ativar o ícone de Modo Turbo no centro da tela (ficando verde), o seu teclado assume atalhos rápidos de transferência:`n")
    $rtb.AppendText("  - F1: Move os arquivos selecionados para o painel oposto.`n")
    $rtb.AppendText("  - F2: Copia os arquivos selecionados para o painel oposto.`n`n")

    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $rtb.AppendText("Barra de Endereços Inteligente`n")
    $rtb.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $rtb.AppendText("• Pressione Ctrl + Backspace na barra de endereços para apagar diretórios inteiros de uma só vez.`n")
    
    $pnlText = New-Object System.Windows.Forms.Panel
    $pnlText.Dock = "Fill"
    $pnlText.Padding = New-Object System.Windows.Forms.Padding(20)
    $pnlText.Controls.Add($rtb)
    $helpForm.Controls.Add($pnlText)

    # ====================================================================
    # TRUQUE NINJA CORRIGIDO: Só aplicamos a regra depois de tudo montado!
    # Lançamos o foco direto para a Janela inteira.
    # ====================================================================
    $rtb.Add_Enter({ $helpForm.Focus() }.GetNewClosure())
    $rtb.Add_GotFocus({ $helpForm.Focus() }.GetNewClosure())

    $helpForm.ShowDialog() | Out-Null
    $helpForm.Dispose()
})

# 2. BOTÃO MINIMIZAR (_) - Posição X: 45
$btnMinApp = New-TitleBtn ([char]0xE921) "#E5E5E5" 10 45
$btnMinApp.Add_Click({ $form.WindowState = 'Minimized' })
$pnlWindowBtns.Controls.Add($btnMinApp)

# 3. BOTÃO MAXIMIZAR ([ ]) - Posição X: 90
$btnMaxApp = New-TitleBtn ([char]0xE922) "#E5E5E5" 10 90
$btnMaxApp.Add_Click({
    if ($form.WindowState -eq 'Maximized') { 
        $form.WindowState = 'Normal'; $btnMaxApp.Text = [char]0xE922 
    } else { 
        # Calcula a área útil do monitor atual pelo Tamanho Máximo
        $form.MaximumSize = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea.Size
        $form.WindowState = 'Maximized'; $btnMaxApp.Text = [char]0xE923 
    }
})
$pnlWindowBtns.Controls.Add($btnMaxApp)

# 4. BOTÃO FECHAR (X) - Posição X: 135
$btnCloseApp = New-TitleBtn ([char]0xE8BB) "#E81123" 10 135
$btnCloseApp.Add_MouseEnter({ $this.ForeColor = "White" })
$btnCloseApp.Add_MouseLeave({ $this.ForeColor = "Black" })
$btnCloseApp.Add_Click({ $form.Close() })
$pnlWindowBtns.Controls.Add($btnCloseApp)

# 3. BACKGROUND CONTAINER (Fica abaixo da barra de titulo)
$bgContainer = New-Object System.Windows.Forms.Panel
$bgContainer.Dock = "Fill"
$bgContainer.BackColor = [System.Drawing.SystemColors]::Control
$form.Controls.Add($bgContainer)
$bgContainer.BringToFront()

# ====================================================================
# 4. RESTAURADOR DE REDIMENSIONAMENTO (BORDAS INVISÍVEIS 3px)
# ====================================================================
function Add-ResizeHandle($Dock, $Cursor, $HitTestCode) {
    $h = New-Object System.Windows.Forms.Panel
    $h.Dock = $Dock
    $h.BackColor = [System.Drawing.Color]::Transparent
    if ($Dock -eq "Bottom" -or $Dock -eq "Top") { $h.Height = 3 } else { $h.Width = 3 }
    $h.Cursor = $Cursor
    $h.Add_MouseDown({
        if ($_.Button -eq 'Left') { [WinDrag]::ReleaseCapture(); [WinDrag]::SendMessage($form.Handle, [WinDrag]::WM_NCLBUTTONDOWN, $HitTestCode, 0) }
    }.GetNewClosure())
    $bgContainer.Controls.Add($h)
    $h.BringToFront()
}

# A CORREÇÃO DOS CURSORES (Envolvidos em Parênteses)
Add-ResizeHandle "Bottom" ([System.Windows.Forms.Cursors]::SizeNS) 15
Add-ResizeHandle "Right" ([System.Windows.Forms.Cursors]::SizeWE) 11
Add-ResizeHandle "Left" ([System.Windows.Forms.Cursors]::SizeWE) 10

# Canto Inferior Direito (Diagonal 12x12px)
$hBR = New-Object System.Windows.Forms.Panel
$hBR.Size = New-Object System.Drawing.Size(12, 12)
$hBR.Cursor = [System.Windows.Forms.Cursors]::SizeNWSE
$hBR.BackColor = [System.Drawing.Color]::Transparent
$hBR.Anchor = "Bottom, Right"
$hBR.Location = New-Object System.Drawing.Point(($bgContainer.Width - 12), ($bgContainer.Height - 12))
$hBR.Add_MouseDown({
    if ($_.Button -eq 'Left') { [WinDrag]::ReleaseCapture(); [WinDrag]::SendMessage($form.Handle, [WinDrag]::WM_NCLBUTTONDOWN, 17, 0) }
})
$bgContainer.Controls.Add($hBR)
$hBR.BringToFront()

# A MESA ORIGINAL VAI DENTRO DO BG CONTAINER
$mainTable = New-Object System.Windows.Forms.TableLayoutPanel
$mainTable.Dock = "Fill"
$mainTable.ColumnCount = 3
$mainTable.RowCount = 1
[void]$mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) 
[void]$mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 50))) 
[void]$mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) 
$bgContainer.Controls.Add($mainTable)
$mainTable.BringToFront()

# --- BLINDAGEM ANTI-ENGASGO 1: Buffer Duplo na Mesa ---
$bindingFlags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
$mainTable.GetType().GetProperty("DoubleBuffered", $bindingFlags).SetValue($mainTable, $true, $null)

# --- BOTÕES CENTRAIS ---
$panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$panelButtons.Dock = "Fill"
$panelButtons.FlowDirection = "TopDown"
$panelButtons.Margin = New-Object System.Windows.Forms.Padding(0)

$mainTable.Controls.Add($panelButtons, 1, 0)

# MÁGICA DE AUTO-CENTRALIZAÇÃO (Atualizada para 7 itens)
$CalcPadding = {
    # 7 itens de 46px (322) + 6 espaços de 20px (120) = 442 pixels de altura total
    $topPadding = [math]::Max(0, [int](($panelButtons.Height - 442) / 2))
    $panelButtons.Padding = New-Object System.Windows.Forms.Padding(2, $topPadding, 0, 0)
}
$panelButtons.Add_Resize($CalcPadding.GetNewClosure())

function New-CmdButton {
    param($SymbolHex, $TooltipText, $Color) 
    $btn = New-Object System.Windows.Forms.Button
    $btn.Size = New-Object System.Drawing.Size(46, 46) 
    $btn.Margin = New-Object System.Windows.Forms.Padding(0, 20, 0, 0) 
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0 
    $btn.BackColor = $Color
    $btn.ForeColor = "White"
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Font = New-Object System.Drawing.Font("Segoe MDL2 Assets", 16)
    $btn.Text = [char][int]$SymbolHex
    
    # CORREÇÃO: Só cria a caixa de texto se a gente enviar um texto!
    if ([string]::IsNullOrWhiteSpace($TooltipText) -eq $false) {
        $tt = New-Object System.Windows.Forms.ToolTip
        $tt.SetToolTip($btn, $TooltipText)
    }
    
    $btn.Add_Click({ Restore-ExplorerFocus })
    return $btn
}

# --- CRIAÇÃO DOS BOTÕES ---

# 1. Trocar Lados
$btnSwap = New-CmdButton -SymbolHex 0xE8AB -TooltipText "Trocar Lados" -Color "#606060"
$btnSwap.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0) 

# 2. Atualizar Pastas
$btnRefresh = New-CmdButton -SymbolHex 0xE72C -TooltipText "Atualizar Pastas" -Color "#009E49"

# 3. INDICADOR VISUAL DE DIREÇÃO (SETA)
$lblDirection = New-Object System.Windows.Forms.Label
$lblDirection.Size = New-Object System.Drawing.Size(46, 46)
$lblDirection.Margin = New-Object System.Windows.Forms.Padding(0, 20, 0, 0) 
$lblDirection.Font = New-Object System.Drawing.Font("Segoe MDL2 Assets", 22, [System.Drawing.FontStyle]::Bold)
$lblDirection.ForeColor = "#0078D7" 
$lblDirection.TextAlign = "MiddleCenter"
$lblDirection.Text = [char]0xF0AF 

$ttDirection = New-Object System.Windows.Forms.ToolTip
$ttDirection.SetToolTip($lblDirection, "A seta aponta para o painel de destino dos arquivos selecionados.")

# 4. Mover
$btnMove = New-CmdButton -SymbolHex 0xE8DE -TooltipText "Mover" -Color "#D83B01"

# 5. Copiar
$btnCopy = New-CmdButton -SymbolHex 0xE8C8 -TooltipText "Copiar" -Color "#0078D7"

# ====================================================================
# 6. NOVO: INTERRUPTOR DO MODO TURBO (VISUAL MINIMALISTA TRANSPARENTE)
# ====================================================================
$global:TurboMode = $false # Começa Desligado

# Em vez de um botão, criamos um Label (Rótulo) com cursor de mãozinha
$lblTurbo = New-Object System.Windows.Forms.Label
$lblTurbo.Size = New-Object System.Drawing.Size(46, 46)
$lblTurbo.Margin = New-Object System.Windows.Forms.Padding(0, 20, 0, 0)
$lblTurbo.Font = New-Object System.Drawing.Font("Segoe MDL2 Assets", 18) # Fonte um pouco maior para destacar
$lblTurbo.TextAlign = "MiddleCenter"
$lblTurbo.Cursor = [System.Windows.Forms.Cursors]::Hand

# VISUAL DESLIGADO (Ícone escuro, fundo transparente)
$lblTurbo.ForeColor = "#606060" # Cinza escuro
$lblTurbo.Text = [char]0xF19E # Slider Vazio

# Caixinha Exclusiva
$ttTurbo = New-Object System.Windows.Forms.ToolTip
$ttTurbo.SetToolTip($lblTurbo, "Modo Turbo (F1=Mover, F2=Copiar) - DESLIGADO")

$lblTurbo.Add_Click({
    $global:TurboMode = -not $global:TurboMode
    if ($global:TurboMode) {
        # VISUAL LIGADO (Ícone colorido)
        $lblTurbo.Text = [char]0xF19F 
        $lblTurbo.ForeColor = "#009E49" # Verde brilhante direto no ícone
        $ttTurbo.SetToolTip($lblTurbo, "Modo Turbo (F1=Mover, F2=Copiar) - LIGADO")
    } else {
        # VOLTA AO DESLIGADO
        $lblTurbo.Text = [char]0xF19E 
        $lblTurbo.ForeColor = "#606060" # Volta pro cinza escuro
        $ttTurbo.SetToolTip($lblTurbo, "Modo Turbo (F1=Mover, F2=Copiar) - DESLIGADO")
    }
    
    # Devolve o foco para o Windows Explorer não travar a navegação
    if (Get-Command "Restore-ExplorerFocus" -ErrorAction SilentlyContinue) { Restore-ExplorerFocus }
}.GetNewClosure())

# Adiciona todos na tela na ordem correta (agora usando o $lblTurbo no final)
$panelButtons.Controls.Add($btnSwap)
$panelButtons.Controls.Add($btnRefresh)
$panelButtons.Controls.Add($lblDirection)
$panelButtons.Controls.Add($btnMove)
$panelButtons.Controls.Add($btnCopy)
$panelButtons.Controls.Add($lblTurbo) # <--- Aqui nós trocamos para a nossa nova Label!

& $CalcPadding

# ====================================================================
# --- CÉREBRO INTELIGENTE DA SETA E TRAVAS ---
# ====================================================================
$directionTimer = New-Object System.Windows.Forms.Timer
$directionTimer.Interval = 200 
$directionTimer.Add_Tick({
    if ($global:ActiveBrowser -ne $null) {
        $isLeftActive = $false
        $isRightActive = $false
        
        if ($global:LeftTabControl -and $global:LeftTabControl.TabPages) {
            foreach ($tab in $global:LeftTabControl.TabPages) {
                if ($tab.Controls.Contains($global:ActiveBrowser)) { $isLeftActive = $true; break }
            }
        }
        
        if ($global:RightTabControl -and $global:RightTabControl.TabPages) {
            foreach ($tab in $global:RightTabControl.TabPages) {
                if ($tab.Controls.Contains($global:ActiveBrowser)) { $isRightActive = $true; break }
            }
        }
        
        if ($isLeftActive) {
            if ($lblDirection.Text -ne [char]0xF0AF) { $lblDirection.Text = [char]0xF0AF }
        } elseif ($isRightActive) {
            if ($lblDirection.Text -ne [char]0xF0B0) { $lblDirection.Text = [char]0xF0B0 }
        }

        try {
            $allowCopy = $true
            $allowMove = $true
            $allowDelete = $true

            # ==========================================================
            # MICRO-FUNÇÃO BLINDADA COM FAXINA DE MEMÓRIA (COM OBJECTS)
            # ==========================================================
            $CheckRestricted = {
                param($b)
                if (-not $b) { return @($false, $false) }
                try {
                    $doc = $b.ActiveXInstance.Document
                    if ($doc) {
                        $isC = $false
                        $isR = $false
                        
                        $folder = $doc.Folder
                        if ($folder) {
                            $self = $folder.Self
                            if ($self) {
                                $fPath = $self.Path
                                $fName = $folder.Title
                                $isC = ($fPath -match "20D04FE0-3AEA-1069-A2D8-08002B30309D" -or $fName -match "Computador")
                                $isR = ($fPath -match "645FF040-5081-101B-9F08-00AA002F954E" -or $fName -match "Lixeira")
                                
                                # Libera os sub-objetos imediatamente
                                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($self) | Out-Null
                            }
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                        }
                        # Libera o documento pai imediatamente
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                        
                        return @($isC, $isR)
                    }
                } catch {}
                return @($false, $false)
            }
            # ==========================================================

            $leftBrowser = if ($global:LeftTabControl -and $global:LeftTabControl.SelectedTab) { $global:LeftTabControl.SelectedTab.Controls[0] } else { $null }
            $rightBrowser = if ($global:RightTabControl -and $global:RightTabControl.SelectedTab) { $global:RightTabControl.SelectedTab.Controls[0] } else { $null }

            $leftStatus = & $CheckRestricted $leftBrowser
            $rightStatus = & $CheckRestricted $rightBrowser

            if ($leftStatus[0] -or $rightStatus[0]) {
                $allowMove = $false
                $allowDelete = $false
            }
            if ($leftStatus[1] -or $rightStatus[1]) {
                $allowCopy = $false
                $allowMove = $false
            }

            if ($btnCopy.Enabled -ne $allowCopy) { 
                $btnCopy.Enabled = $allowCopy
                $btnCopy.BackColor = if ($allowCopy) { "#0078D7" } else { "#A0A0A0" }
            }
            if ($btnMove.Enabled -ne $allowMove) { 
                $btnMove.Enabled = $allowMove
                $btnMove.BackColor = if ($allowMove) { "#D83B01" } else { "#A0A0A0" }
            }
        } catch {}
    }
})
$directionTimer.Start()

$btnSwap.Add_Click({
    try {
        $leftTabs = @()
        foreach ($tab in $global:LeftTabControl.TabPages) { $leftTabs += $tab }
        
        $rightTabs = @()
        foreach ($tab in $global:RightTabControl.TabPages) { $rightTabs += $tab }

        $global:LeftTabControl.TabPages.Clear()
        $global:RightTabControl.TabPages.Clear()

        foreach ($tab in $rightTabs) { $global:LeftTabControl.TabPages.Add($tab) }
        foreach ($tab in $leftTabs)  { $global:RightTabControl.TabPages.Add($tab) }
    } catch {}
}.GetNewClosure())

# ==============================================================================
# --- TRADUTOR C# PARA TRANSFERÊNCIA NATIVA DO WINDOWS ---
# ==============================================================================
$FileOpCSharp = @"
using System;
using System.Runtime.InteropServices;

public class Win32FileOp {
    public enum FO_FUNC : uint { FO_MOVE = 0x0001, FO_COPY = 0x0002 }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SHFILEOPSTRUCT {
        public IntPtr hwnd;
        public FO_FUNC wFunc;
        [MarshalAs(UnmanagedType.LPWStr)] public string pFrom;
        [MarshalAs(UnmanagedType.LPWStr)] public string pTo;
        public ushort fFlags;
        public bool fAnyOperationsAborted;
        public IntPtr hNameMappings;
        [MarshalAs(UnmanagedType.LPWStr)] public string lpszProgressTitle;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int SHFileOperation(ref SHFILEOPSTRUCT fileop);

    public static void DoOperation(string[] sources, string dest, bool isMove) {
        if (sources == null || sources.Length == 0) return;
        SHFILEOPSTRUCT shf = new SHFILEOPSTRUCT();
        shf.wFunc = isMove ? FO_FUNC.FO_MOVE : FO_FUNC.FO_COPY;
        
        shf.pFrom = string.Join("\0", sources) + "\0\0";
        shf.pTo = dest + "\0\0";
        shf.fFlags = 0x0200; 
        
        SHFileOperation(ref shf);
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]'Win32FileOp').Type) {
    Add-Type -TypeDefinition $FileOpCSharp
}

# ==============================================================================
# --- MOTOR DE TRANSFERÊNCIA DE ARQUIVOS ---
# ==============================================================================

$PerformFileAction = {
    param([bool]$IsMove)

    if ($global:ActiveBrowser -eq $null) { return }

    $shellView = $global:ActiveBrowser.ActiveXInstance.Document
    if (-not $shellView) { return }

    $selectedItems = $shellView.SelectedItems()
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Selecione pelo menos um arquivo ou pasta para transferir.", "Aviso", 0, [System.Windows.Forms.MessageBoxIcon]::Warning)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selectedItems) | Out-Null # Limpa se estiver vazio
        return
    }

    $isLeftActive = $false
    foreach ($tab in $global:LeftTabControl.TabPages) {
        if ($tab.Controls.Contains($global:ActiveBrowser)) { $isLeftActive = $true; break }
    }

    $targetTabControl = if ($isLeftActive) { $global:RightTabControl } else { $global:LeftTabControl }
    if ($targetTabControl.SelectedTab -eq $null -or $targetTabControl.SelectedTab.Controls.Count -eq 0) { 
        if ($selectedItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selectedItems) | Out-Null }
        if ($shellView) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null }
        return 
    }

    $targetBrowser = $targetTabControl.SelectedTab.Controls[0]
    if (-not $targetBrowser.Url) { 
        if ($selectedItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selectedItems) | Out-Null }
        if ($shellView) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null }
        return 
    }

    $sourcePath = $global:ActiveBrowser.Url.LocalPath
    $targetPath = $targetBrowser.Url.LocalPath

    if ($sourcePath -eq $targetPath) { 
        if ($selectedItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selectedItems) | Out-Null }
        if ($shellView) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null }
        return 
    }

    $sourcePaths = @()
    foreach ($item in $selectedItems) {
        if ($item.Path) { $sourcePaths += $item.Path }
    }

    try {
        [Win32FileOp]::DoOperation($sourcePaths, $targetPath, $IsMove)
        
        $global:ActiveBrowser.Refresh()
        $targetBrowser.Refresh()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erro ao transferir: $_", "Erro", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    # --- CORREÇÃO DO MICRO-VAZAMENTO ---
    if ($selectedItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selectedItems) | Out-Null }
    if ($shellView) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null }
}

$btnCopy.Add_Click({ & $PerformFileAction -IsMove $false })
$btnMove.Add_Click({ & $PerformFileAction -IsMove $true })
$btnRefresh.Add_Click({ 
    if ($global:ActiveBrowser) { $global:ActiveBrowser.Refresh() }
    if ($global:LeftTabControl.SelectedTab) { $global:LeftTabControl.SelectedTab.Controls[0].Refresh() }
    if ($global:RightTabControl.SelectedTab) { $global:RightTabControl.SelectedTab.Controls[0].Refresh() }
})

# ====================================================================
# --- MOTOR INVISÍVEL DE TECLADO (GUARDA-COSTAS DO MODO TURBO) ---
# ====================================================================
if (-not ("MessageFilter" -as [type])) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Windows.Forms;
    public class MessageFilter : IMessageFilter {
        public delegate bool KeyHandler(Keys key);
        public KeyHandler OnKeyDown;
        
        public bool PreFilterMessage(ref Message m) {
            // Intercepta apenas a mensagem WM_KEYDOWN (0x0100)
            if (m.Msg == 0x0100) { 
                Keys key = (Keys)m.WParam.ToInt32();
                if (OnKeyDown != null) {
                    return OnKeyDown(key); // Se retornar true, mata a tecla aqui!
                }
            }
            return false;
        }
    }
"@ -ReferencedAssemblies "System.Windows.Forms" # <--- AQUI ESTÁ A CORREÇÃO MAGICA
}

$global:MyKeyFilter = New-Object MessageFilter
$global:MyKeyFilter.OnKeyDown = {
    param([System.Windows.Forms.Keys]$key)
    
    # Se o Turbo não estiver ligado, deixa a tecla passar
    if (-not $global:TurboMode) { return $false }

    if ($key -eq [System.Windows.Forms.Keys]::F1) {
        if ($btnMove.Enabled) { $btnMove.PerformClick() }
        return $true # MATA A TECLA (Bloqueia a Ajuda)
    }
    if ($key -eq [System.Windows.Forms.Keys]::F2) {
        if ($btnCopy.Enabled) { $btnCopy.PerformClick() }
        return $true # MATA A TECLA (Bloqueia o Renomear)
    }
    
    return $false
}
[System.Windows.Forms.Application]::AddMessageFilter($global:MyKeyFilter)

# Limpeza vitalícia ao fechar o programa
$form.Add_FormClosed({
    [System.Windows.Forms.Application]::RemoveMessageFilter($global:MyKeyFilter)
    
    # ====================================================================
    # --- FECHA O ESQUADRÃO TÁTICO AO FECHAR A JANELA ---
    # Impede que as threads fiquem rodando como zumbis ocultos na memória
    # ====================================================================
    if ($global:RunspacePool) {
        try {
            $global:RunspacePool.Close()
            $global:RunspacePool.Dispose()
        } catch {}
    }
})

# ==============================================================================
# --- EXTRATOR DE ÍCONES NATIVOS DO WINDOWS (SHELL32) ---
# ==============================================================================
if (-not ("IconExtractor" -as [type])) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class IconExtractor {
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool DestroyIcon(IntPtr handle);
    }
"@
}

# --- NAVEGADOR COM ABAS (CÉREBRO INDEPENDENTE POR ABA) ---
function New-BrowserPane {
    # CORREÇÃO: Removemos o [string[]] e trocamos por [array] para aceitar os Objetos do Recovery!
    param ($TableControl, $ColumnIndex, [array]$InitialPaths)

    $panel = New-Object System.Windows.Forms.Panel; $panel.Dock = "Fill"; $TableControl.Controls.Add($panel, $ColumnIndex, 0)

    # 1. STATUS (AGORA UM PAINEL DIVIDIDO E ALINHADO)
    $statusPanel = New-Object System.Windows.Forms.Panel; $statusPanel.Dock = "Bottom"; $statusPanel.Height = 24; $statusPanel.BackColor = "#F5F5F5"; $statusPanel.BorderStyle="FixedSingle"; $panel.Controls.Add($statusPanel)
    
    # O SEGREDO DO ALINHAMENTO: MinimumSize = (0, 24) força a altura a ser igual à do painel!
    $lblDisk = New-Object System.Windows.Forms.Label; $lblDisk.Dock = "Right"; $lblDisk.TextAlign="MiddleRight"; $lblDisk.AutoSize=$true; $lblDisk.MinimumSize = New-Object System.Drawing.Size(0, 24); $lblDisk.Padding = New-Object System.Windows.Forms.Padding(0, 0, 5, 0); $lblDisk.BackColor = [System.Drawing.Color]::Transparent; $statusPanel.Controls.Add($lblDisk)
    
    $lblStatus = New-Object System.Windows.Forms.Label; $lblStatus.Dock = "Fill"; $lblStatus.TextAlign="MiddleLeft"; $lblStatus.AutoSize=$false; $lblStatus.AutoEllipsis=$true; $lblStatus.BackColor = [System.Drawing.Color]::Transparent; $statusPanel.Controls.Add($lblStatus)
    $lblStatus.BringToFront()

    # 2. CONTAINER DE TOPO
    $topContainer = New-Object System.Windows.Forms.Panel; $topContainer.Dock = "Top"; $topContainer.AutoSize = $true; $panel.Controls.Add($topContainer)

    # ====================================================================
    # O SEGREDO: Criar o TabControl ANTES para que os botões saibam que ele existe!
    # ====================================================================
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.DrawMode = "OwnerDrawFixed"
    $tabControl.Padding = New-Object System.Drawing.Point(12, 4)
    $panel.Controls.Add($tabControl)
    $tabControl.BringToFront()

    # ====================================================================
    # 2.1 MARCADORES, ATALHOS E EXIBIÇÃO (NO TOPO)
    # ====================================================================
    $bookmarksBar = New-Object System.Windows.Forms.Panel; $bookmarksBar.Dock = "Top"; $bookmarksBar.Height = 26; $bookmarksBar.BackColor = "#EEEEEE"; $topContainer.Controls.Add($bookmarksBar)
    
    # --- NOSSA ADIÇÃO: MICRO-FUNÇÃO DO ESPAÇADOR ---
    function Add-Spacer {
        $spc = New-Object System.Windows.Forms.Label
        $spc.AutoSize = $false; $spc.Width = 1; $spc.Dock = "Left"
        $spc.BackColor = [System.Drawing.Color]::Transparent
        $bookmarksBar.Controls.Add($spc); $spc.BringToFront()
    }
    
    # ====================================================================
    # A SUA "CHAVE": Controle de estado exato e instantâneo
    # ====================================================================
    $menuKey = @{ Bookmarks = $false; Shortcuts = $false; View = $false; Sort = $false }

    # --- BOTÃO MARCADORES ---
    $btnBookmarks = New-Object System.Windows.Forms.Button
    $btnBookmarks.Text = "Marcadores"
    $btnBookmarks.Dock = "Left"
    
    # A MÁGICA ELÁSTICA
    $btnBookmarks.AutoSize = $true 
    $btnBookmarks.AutoSizeMode = "GrowAndShrink"

    # 3. Estilo Standard e Alinhamento centralizado
    $btnBookmarks.FlatStyle = "Popup" 
    $btnBookmarks.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    
    # 4. Remove o colchão de margens
    $btnBookmarks.Padding = New-Object System.Windows.Forms.Padding(0)
    $btnBookmarks.Margin = New-Object System.Windows.Forms.Padding(0)
    
    $bookmarksBar.Controls.Add($btnBookmarks)
    $btnBookmarks.BringToFront()
    
    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    # Evento de fechamento: Deixa a chave se o mouse causou o fechamento
    $ctxMenu.Add_Closed({ 
        $mousePos = $btnBookmarks.PointToClient([System.Windows.Forms.Cursor]::Position)
        if ($btnBookmarks.ClientRectangle.Contains($mousePos)) { $menuKey.Bookmarks = $true }
        
        # O TRUQUE NINJA: Tira o foco do botão e joga para a barra de fundo
        Restore-ExplorerFocus
    }.GetNewClosure())
    
    $btnBookmarks.Add_Click({
        # Se a chave estiver presente, consome a chave e não abre o menu
        if ($menuKey.Bookmarks) { $menuKey.Bookmarks = $false; return }
        
        # --- PREVENÇÃO DE VAZAMENTO: Destrói os itens antigos antes de recriar ---
        while ($ctxMenu.Items.Count -gt 0) { $ctxMenu.Items[0].Dispose() }
        
        $itemMgr = $ctxMenu.Items.Add("Gerenciador de Marcadores"); $itemMgr.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $itemMgr.Add_Click({ Show-BookmarkManager })
        [void]$ctxMenu.Items.Add("-")
        $data = Get-BookmarksData; if ($data -and $data.Children) { Build-ContextMenu $ctxMenu.Items $data.Children }
        $ctxMenu.Show($btnBookmarks, 0, $btnBookmarks.Height)
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

    # --- NOVO BOTÃO: ATALHOS UNIVERSAIS (MOVIDO PARA O MEIO) ---
    $btnShortcuts = New-Object System.Windows.Forms.Button
    $btnShortcuts.Text = "Atalhos"
    $btnShortcuts.Dock = "Left"
    
    # A MÁGICA ELÁSTICA
    $btnShortcuts.AutoSize = $true 
    $btnShortcuts.AutoSizeMode = "GrowAndShrink"

    # 3. Estilo Standard e Alinhamento centralizado
    $btnShortcuts.FlatStyle = "Popup" 
    $btnShortcuts.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    
    # 4. Remove o colchão de margens
    $btnShortcuts.Padding = New-Object System.Windows.Forms.Padding(0)
    $btnShortcuts.Margin = New-Object System.Windows.Forms.Padding(0)
    
    $bookmarksBar.Controls.Add($btnShortcuts)
    $btnShortcuts.BringToFront()

    $ctxShortcutsMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $ctxShortcutsMenu.Add_Closed({ 
        $mousePos = $btnShortcuts.PointToClient([System.Windows.Forms.Cursor]::Position)
        if ($btnShortcuts.ClientRectangle.Contains($mousePos)) { $menuKey.Shortcuts = $true }
        
        # Remove o foco (destaque) do botão assim que o menu fechar
        Restore-ExplorerFocus
    }.GetNewClosure())
    
    $shortcutsList = @(
        @{ Name="Meu Computador";   Path="shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"; Dll="imageres.dll"; Icon= -109 },
        @{ Name="Lixeira";          Path="shell:::{645FF040-5081-101B-9F08-00AA002F954E}"; Dll="imageres.dll"; Icon= -54 },
        @{ Name="Área de Trabalho"; Path=[Environment]::GetFolderPath('Desktop');          Dll="imageres.dll"; Icon= -110 },
        @{ Name="Downloads";        Path="$env:USERPROFILE\Downloads";                     Dll="imageres.dll"; Icon= -184 },
        @{ Name="Documentos";       Path=[Environment]::GetFolderPath('MyDocuments');      Dll="imageres.dll"; Icon= -112 },
        @{ Name="Imagens";          Path=[Environment]::GetFolderPath('MyPictures');       Dll="imageres.dll"; Icon= -113 },
        @{ Name="Músicas";          Path=[Environment]::GetFolderPath('MyMusic');          Dll="imageres.dll"; Icon= -108 },
        @{ Name="Vídeos";           Path="$env:USERPROFILE\Vídeos";                        Dll="imageres.dll"; Icon= -189 }
    )

    foreach ($sc in $shortcutsList) {
        $item = $ctxShortcutsMenu.Items.Add($sc.Name)
        $item.Tag = $sc.Path
        
        $dllPath = "$env:windir\System32\$($sc.Dll)"
        
        $hIcon = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllPath, $sc.Icon)
        if ($hIcon -ne [IntPtr]::Zero) {
            $ico = [System.Drawing.Icon]::FromHandle($hIcon)
            $item.Image = $ico.ToBitmap()
            
            # --- CORREÇÃO DE VAZAMENTO: Destrói o ícone após virar Bitmap ---
            $ico.Dispose()
            [IconExtractor]::DestroyIcon($hIcon) | Out-Null
        }

        $item.Add_Click({
            param($sender, $e)
            if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
                $b = $tabControl.SelectedTab.Controls[0]
                $b.Navigate($sender.Tag) 
            }
        }.GetNewClosure())
    }

    $btnShortcuts.Add_Click({
        if ($menuKey.Shortcuts) { $menuKey.Shortcuts = $false; return }
        $ctxShortcutsMenu.Show($btnShortcuts, 0, $btnShortcuts.Height)
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

# --- BOTAO Exibição (MOVIDO PARA A DIREITA DOS ATALHOS) ---
    $btnViewMode = New-Object System.Windows.Forms.Button
    $btnViewMode.Text = "Exibição" 
    $btnViewMode.Dock = "Left"
    
    # A MÁGICA ELÁSTICA
    $btnViewMode.AutoSize = $true 
    $btnViewMode.AutoSizeMode = "GrowAndShrink"

    # 3. Estilo Standard e Alinhamento centralizado
    $btnViewMode.FlatStyle = "Popup" 
    $btnViewMode.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    
    # 4. Remove o colchão de margens
    $btnViewMode.Padding = New-Object System.Windows.Forms.Padding(0)
    $btnViewMode.Margin = New-Object System.Windows.Forms.Padding(0)
    
    $bookmarksBar.Controls.Add($btnViewMode)
    $btnViewMode.BringToFront() 

    $ctxViewMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $ctxViewMenu.AutoSize = $false 
    
    $ctxViewMenu.Add_Closed({ 
        $mousePos = $btnViewMode.PointToClient([System.Windows.Forms.Cursor]::Position)
        if ($btnViewMode.ClientRectangle.Contains($mousePos)) { $menuKey.View = $true }
        
        # Limpa o destaque visual do botão Exibição
        Restore-ExplorerFocus
    }.GetNewClosure())

    $viewOptions = @("Ícones Extra Grandes", "Ícones Grandes", "Ícones Medios", "Ícones Pequenos", "Lista", "Detalhes", "Lado a Lado", "Conteúdo")
    
    foreach ($opt in $viewOptions) {
        $menuItem = $ctxViewMenu.Items.Add($opt)
        $menuItem.Add_Click({
            param($sender, $e)
            $newView = $sender.Text
            if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
                $activeBrowser = $tabControl.SelectedTab.Controls[0]
                $activeBrowser.Tag.ViewMode = $newView
                Set-View-Mode-Scanner -Browser $activeBrowser -ViewName $newView
            }
        }.GetNewClosure())
    }

    $ctxViewMenu.Add_Opening({
        $totalHeight = $ctxViewMenu.Padding.Top + $ctxViewMenu.Padding.Bottom
        foreach ($item in $ctxViewMenu.Items) { $totalHeight += $item.Height }
        
        $ctxViewMenu.Size = New-Object System.Drawing.Size(190, ($totalHeight + 5))

        $currentView = "Detalhes" 
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $activeBrowser = $tabControl.SelectedTab.Controls[0]
            
            # --- A CIRURGIA: Leitura em tempo real do motor do Windows ---
            try {
                $viewDoc = $activeBrowser.ActiveXInstance.Document
                if ($viewDoc) {
                    $vMode = $viewDoc.CurrentViewMode
                    $iSize = $viewDoc.IconSize

                    # Traduz os códigos nativos do Windows para os nomes do nosso menu
                    switch ($vMode) {
                        1 {
                            if ($iSize -ge 256) { $currentView = "Ícones Extra Grandes" }
                            elseif ($iSize -ge 96) { $currentView = "Ícones Grandes" }
                            else { $currentView = "Ícones Medios" }
                        }
                        2 { $currentView = "Ícones Pequenos" }
                        3 { $currentView = "Lista" }
                        4 { $currentView = "Detalhes" }
                        5 { $currentView = "Miniaturas" }
                        6 { $currentView = "Lado a Lado" }
                        8 { $currentView = "Conteúdo" }
                    }

                    # Atualiza a memória do script para ficar igual à realidade do ecrã
                    if ($activeBrowser.Tag) { $activeBrowser.Tag.ViewMode = $currentView }

                    # Limpeza imediata da memória (Blindagem Anti-Leak)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($viewDoc) | Out-Null
                }
            } catch {
                # Se falhar (por a pasta estar a carregar), usa a memória como plano de segurança
                if ($activeBrowser.Tag -and $activeBrowser.Tag.ViewMode) { $currentView = $activeBrowser.Tag.ViewMode }
            }
        }
        
        foreach ($item in $ctxViewMenu.Items) {
            if ($item.Text -eq $currentView) { 
                $item.Checked = $true 
            } else { 
                $item.Checked = $false 
            }
        }
    }.GetNewClosure())

    $btnViewMode.Add_Click({ 
        if ($menuKey.View) { $menuKey.View = $false; return }
        $ctxViewMenu.Show($btnViewMode, 0, $btnViewMode.Height) 
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

# --- BOTAO CLASSIFICAR (A DIREITA DE Exibição) ---
    $btnSort = New-Object System.Windows.Forms.Button
    $btnSort.Text = "Classificar"
    $btnSort.Dock = "Left"
    
    # A MÁGICA ELÁSTICA
    $btnSort.AutoSize = $true 
    $btnSort.AutoSizeMode = "GrowAndShrink"

    $btnSort.FlatStyle = "Popup" 
    $btnSort.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $btnSort.Padding = New-Object System.Windows.Forms.Padding(0)
    $btnSort.Margin = New-Object System.Windows.Forms.Padding(0)
    
    $bookmarksBar.Controls.Add($btnSort)
    $btnSort.BringToFront() 

    $ctxSortMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $ctxSortMenu.Add_Closed({ 
        $mousePos = $btnSort.PointToClient([System.Windows.Forms.Cursor]::Position)
        if ($btnSort.ClientRectangle.Contains($mousePos)) { $menuKey.Sort = $true }
        
        # O toque final: tira o foco do botão Classificar
        Restore-ExplorerFocus
    }.GetNewClosure())

    $sortOptions = @(
        @{ Label = "Nome"; Prop = "System.ItemNameDisplay" },
        @{ Label = "Data de modificação"; Prop = "System.DateModified" },
        @{ Label = "Tipo"; Prop = "System.ItemTypeText" },
        @{ Label = "Tamanho"; Prop = "System.Size" }
    )

    foreach ($opt in $sortOptions) {
        $menuItem = $ctxSortMenu.Items.Add($opt.Label)
        $menuItem.Tag = $opt.Prop
        $menuItem.Add_Click({
            param($sender, $e)
            if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
                $b = $tabControl.SelectedTab.Controls[0]
                try {
                    $view = $b.ActiveXInstance.Document
                    $currentSort = $view.SortColumns
                    $newProp = $sender.Tag
                    
                    if ($currentSort -match $newProp) {
                        # Se clicar na MESMA categoria, inverte (Alterna a direcao)
                        $prefix = if ($currentSort.StartsWith("prop:-")) { "prop:" } else { "prop:-" }
                    } else {
                        # NOVA CATEGORIA: Forca sempre Crescente
                        # O Windows inverte o padrao de Data e Tamanho, entao enviamos prop:- para forcar o Crescente neles
                        if ($newProp -match "Size|DateModified") {
                            $prefix = "prop:-" 
                        } else {
                            $prefix = "prop:"
                        }
                    }
                    
                    $view.SortColumns = "$prefix$newProp;"
                    
                    # --- FAXINA COM OBJECT ---
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
                } catch {}
            }
        }.GetNewClosure())
    }

    [void]$ctxSortMenu.Items.Add("-")

    $menuAsc = $ctxSortMenu.Items.Add("Crescente")
    $menuAsc.Add_Click({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]
            try {
                $view = $b.ActiveXInstance.Document
                $currentSort = $view.SortColumns
                if ($currentSort -match "prop:([+-]?)(.*)") {
                    $propName = $matches[2] -replace ";", ""
                    # Se for Data ou Tamanho, Crescente exige prop:-
                    $prefix = if ($propName -match "Size|DateModified") { "prop:-" } else { "prop:" }
                    $view.SortColumns = "$prefix$propName;"
                }
                
                # --- FAXINA COM OBJECT ---
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
            } catch {}
        }
    }.GetNewClosure())

    $menuDesc = $ctxSortMenu.Items.Add("Decrescente")
    $menuDesc.Add_Click({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]
            try {
                $view = $b.ActiveXInstance.Document
                $currentSort = $view.SortColumns
                if ($currentSort -match "prop:([+-]?)(.*)") {
                    $propName = $matches[2] -replace ";", ""
                    # Se for Data ou Tamanho, Decrescente exige prop:
                    $prefix = if ($propName -match "Size|DateModified") { "prop:" } else { "prop:-" }
                    $view.SortColumns = "$prefix$propName;"
                }
                
                # --- FAXINA COM OBJECT ---
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
            } catch {}
        }
    }.GetNewClosure())

    $ctxSortMenu.Add_Opening({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]
            try {
                $view = $b.ActiveXInstance.Document
                $currentSort = $view.SortColumns
                
                # Marca a categoria atual
                for ($i = 0; $i -lt 4; $i++) {
                    $item = $ctxSortMenu.Items[$i]
                    if ($currentSort -match $item.Tag) { 
                        $item.Checked = $true 
                    } else { 
                        $item.Checked = $false 
                    }
                }
                
                # Leitura inteligente da direcao para marcar Crescente ou Decrescente
                $hasMinus = $currentSort.StartsWith("prop:-")
                $isDateOrSize = ($currentSort -match "Size|DateModified")
                
                if ($isDateOrSize) {
                    # Se for Tamanho/Data, o traco (-) significa Crescente
                    $isDesc = -not $hasMinus
                } else {
                    # Se for Nome/Tipo, o traco (-) significa Decrescente
                    $isDesc = $hasMinus
                }

                $ctxSortMenu.Items[6].Checked = $isDesc         # Decrescente (Indice 6)
                $ctxSortMenu.Items[5].Checked = (-not $isDesc)  # Crescente (Indice 5)
                
                # --- FAXINA COM OBJECT ---
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
            } catch {}
        }
    }.GetNewClosure())

    $btnSort.Add_Click({ 
        if ($menuKey.Sort) { $menuKey.Sort = $false; return }
        $ctxSortMenu.Show($btnSort, 0, $btnSort.Height) 
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

# ====================================================================
# --- BOTAO PREVIEW E JANELA FLUTUANTE DEDICADA ---
# ====================================================================
    # 1. Compilador do Windows Media Player para o PowerShell (MANTIDO)
    $wmpCode = @'
    using System.Windows.Forms;
    public class WMPHost : AxHost {
        public WMPHost() : base("6BF52A52-394A-11d3-B153-00C04F79FAA6") {}
        public object GetMediaPlayer() { return this.GetOcx(); }
    }
'@
    if (-not ([System.Management.Automation.PSTypeName]'WMPHost').Type) {
        try { Add-Type -TypeDefinition $wmpCode -ReferencedAssemblies "System.Windows.Forms" -ErrorAction SilentlyContinue } catch {}
    }

    # 2. Compilador Extrator de Ícones (Puxando a versão 16x16 perfeita)
    $extractorCode = @'
    using System;
    using System.Drawing;
    using System.Runtime.InteropServices;
    public class DllIconExtractor {
        [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
        public static extern int ExtractIconEx(string szFileName, int nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, uint nIcons);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyIcon(IntPtr hIcon);

        public static Bitmap GetIcon(string path, int index) {
            IntPtr large = IntPtr.Zero;
            IntPtr small = IntPtr.Zero;
            ExtractIconEx(path, index, out large, out small, 1);
            
            // PEGA O ICONE PEQUENO (16x16) para nao estourar o botao
            if (small != IntPtr.Zero) {
                Icon icon = Icon.FromHandle(small);
                Bitmap bmp = icon.ToBitmap();
                DestroyIcon(small);
                if (large != IntPtr.Zero) DestroyIcon(large);
                return bmp;
            }
            return null;
        }
    }
'@
    if (-not ([System.Management.Automation.PSTypeName]'DllIconExtractor').Type) {
        try { Add-Type -TypeDefinition $extractorCode -ReferencedAssemblies System.Drawing, System.Windows.Forms -ErrorAction SilentlyContinue } catch {}
    }

    # --- BOTAO PREVIEW (Somente Icone 16x16) ---
    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Dock = "Left"

    try {
        $dllPath = "$env:SystemRoot\System32\mshtml.dll"
        $btnPreview.Image = [DllIconExtractor]::GetIcon($dllPath, 6)
    } catch {}

    $btnPreview.Text = ""
    $btnPreview.AutoSize = $false
    $btnPreview.Width = 30  # Largura ajustada para abraçar o icone 16x16
    
    # ESTILO ATUALIZADO
    $btnPreview.FlatStyle = "Popup" 
    
    $btnPreview.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $btnPreview.Padding = New-Object System.Windows.Forms.Padding(0)
    $btnPreview.Margin = New-Object System.Windows.Forms.Padding(0)

    # Dica flutuante para manter a usabilidade
    $ttPreview = New-Object System.Windows.Forms.ToolTip
    $ttPreview.SetToolTip($btnPreview, "Painel de Preview")

    # TRUQUE NINJA: Limpa o destaque visual logo após o clique
    $btnPreview.Add_Click({
        Restore-ExplorerFocus
    }.GetNewClosure())

    $bookmarksBar.Controls.Add($btnPreview)
    $btnPreview.BringToFront()

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

    # --- NOVO: BOTAO PROPRIEDADES (Usando IconExtractor com Bloqueio Inteligente) ---
    $btnProp = New-Object System.Windows.Forms.Button
    $btnProp.Dock = "Left"
    $btnProp.Width = 26
    
    # ESTILO ATUALIZADO PARA POPUP
    $btnProp.FlatStyle = "Popup" 
    
    $btnProp.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnProp.Text = ""
    $btnProp.Enabled = $false # INICIA BLOQUEADO (Padrao)
    $bookmarksBar.Controls.Add($btnProp)
    $btnProp.BringToFront()

    $ttProp = New-Object System.Windows.Forms.ToolTip
    $ttProp.SetToolTip($btnProp, "Propriedades")

    $dllMmc = "$env:SystemRoot\System32\prnfldr.dll"
    try {
        # Extrai o icone 6 da nova DLL
        $hIconProp = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllMmc, 6)
        if ($hIconProp -ne [IntPtr]::Zero) {
            $icoProp = [System.Drawing.Icon]::FromHandle($hIconProp)
            $origBmpProp = $icoProp.ToBitmap()
            
            $newSizeProp = New-Object System.Drawing.Size(16, 16)
            $btnProp.Image = New-Object System.Drawing.Bitmap($origBmpProp, $newSizeProp)
            
            # Limpeza de memoria garantida
            $origBmpProp.Dispose()
            $icoProp.Dispose()
            [IconExtractor]::DestroyIcon($hIconProp) | Out-Null
        } else { 
            $btnProp.Text = "P" 
        }
    } catch { 
        $btnProp.Text = "P" 
    }

    # Acao NATIVA de Propriedades
    $btnProp.Add_Click({
        if ($tabControl.SelectedTab -ne $null -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $currBrowser = $tabControl.SelectedTab.Controls[0]
            $shellView = $currBrowser.ActiveXInstance.Document
            if ($shellView) {
                $folder = $shellView.Folder
                if ($folder) {
                    $self = $folder.Self
                    if ($self) {
                        $selItems = $shellView.SelectedItems()
                        
                        if ($selItems.Count -gt 0) {
                            foreach ($item in $selItems) { try { $item.InvokeVerb("properties") } catch {} }
                        } else {
                            try { $self.InvokeVerb("properties") } catch {}
                        }

                        if ($selItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selItems) | Out-Null }
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($self) | Out-Null
                    }
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null
            }
        }
        Restore-ExplorerFocus
    }.GetNewClosure())

    # ====================================================================
    # --- A PARTIR DAQUI COMEÇA A JANELA FLUTUANTE DE PREVIEW ---
    # ====================================================================

    $sideName = if ($ColumnIndex -eq 0) { "Esquerdo" } else { "Direito" }
    
    $previewForm = New-Object System.Windows.Forms.Form
    $previewForm.Text = "Preview ($sideName)"
    $previewForm.Size = New-Object System.Drawing.Size(450, 450)
    $previewForm.StartPosition = "Manual"
    $previewForm.ShowInTaskbar = $true 
    
    # --- ALTERAÇÃO: Muda o estilo da janela para liberar os botões ---
    $previewForm.FormBorderStyle = "Sizable"
    $previewForm.MaximizeBox = $true
    $previewForm.MinimizeBox = $true
    $previewForm.ShowIcon = $false

    # ====================================================================
    # --- SISTEMA DE MEMÓRIA DUPLA E EFEITO ELÁSTICO (ANTI-DESSINCRONIA) ---
    # ====================================================================
    $previewMem = @{ A_Loc = $null; A_Size = $null; A_Max = $false; B_Loc = $null; B_Size = $null; B_Max = $false; Active = "A" }

    $SwitchProfile = {
        param([string]$NewProfile)
        if ($previewMem.Active -eq $NewProfile -and $previewForm.Visible) { return } 
        
        $isMax = ($previewForm.WindowState -eq 'Maximized')

        # 1. Salva o estado atual (Ignora tamanho se estiver maximizado)
        if ($previewMem.Active -eq "A") {
            if (-not $isMax) { $previewMem.A_Loc = $previewForm.Location; $previewMem.A_Size = $previewForm.Size }
            $previewMem.A_Max = $isMax
        } else {
            $scr = [System.Windows.Forms.Screen]::FromControl($previewForm)
            if ($scr.Primary) { 
                if (-not $isMax) { $previewMem.B_Loc = $previewForm.Location; $previewMem.B_Size = $previewForm.Size }
                $previewMem.B_Max = $isMax
            }
        }

        # 1.5 Desativa o maximizado antes da transição para evitar distorções
        $previewForm.WindowState = 'Normal'

        # 2. Muda para o novo perfil
        $previewMem.Active = $NewProfile

        if ($NewProfile -eq "B") {
            if ($previewMem.B_Size -ne $null) { $previewForm.Size = $previewMem.B_Size }
            if ($previewMem.B_Loc -ne $null) { 
                $previewForm.Location = $previewMem.B_Loc 
            } else {
                $prim = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
                $previewForm.Location = New-Object System.Drawing.Point(($prim.X + 150), ($prim.Y + 150))
            }
            
            $scr = [System.Windows.Forms.Screen]::FromControl($previewForm)
            if (-not $scr.Primary) {
                $prim = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
                $previewForm.Location = New-Object System.Drawing.Point(($prim.X + 150), ($prim.Y + 150))
            }
            
            # Restaura a tela cheia se era assim que estava antes
            if ($previewMem.B_Max) { $previewForm.WindowState = 'Maximized' }
            
        } else {
            if ($previewMem.A_Size -ne $null) { $previewForm.Size = $previewMem.A_Size }
            if ($previewMem.A_Loc -ne $null) { $previewForm.Location = $previewMem.A_Loc }
            
            # Restaura a tela cheia se era assim que estava antes
            if ($previewMem.A_Max) { $previewForm.WindowState = 'Maximized' }
        }
    }.GetNewClosure()

    $previewForm.Add_ResizeEnd({
        if ($previewMem.Active -eq "B") {
            $scr = [System.Windows.Forms.Screen]::FromControl($previewForm)
            if (-not $scr.Primary) {
                $prim = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
                if ($previewMem.B_Loc -ne $null) { $previewForm.Location = $previewMem.B_Loc }
                else { $previewForm.Location = New-Object System.Drawing.Point(($prim.X + 150), ($prim.Y + 150)) }
                
                if ($wmpPreview -ne $null -and $wmpPreview.Visible) {
                    try {
                        $ocx = $wmpPreview.GetMediaPlayer()
                        $pos = $ocx.controls.currentPosition
                        $ocx.controls.stop(); $ocx.controls.play(); $ocx.controls.currentPosition = $pos
                    } catch {}
                }
            } else {
                # Atualiza a memória B (Apenas se NÃO estiver Maximizada)
                if ($previewForm.WindowState -ne 'Maximized') {
                    $previewMem.B_Loc = $previewForm.Location
                    $previewMem.B_Size = $previewForm.Size
                }
            }
        } else {
            # Salva livremente no Perfil A (Apenas se NÃO estiver Maximizada)
            if ($previewForm.WindowState -ne 'Maximized') {
                $previewMem.A_Loc = $previewForm.Location
                $previewMem.A_Size = $previewForm.Size
            }
        }
    }.GetNewClosure())

    # ====================================================================
    # --- A CAIXA PRETA: BLINDAGEM CONTRA O WINDOWS (AMNÉSIA DE MAXIMIZAR) ---
    # ====================================================================
    $previewForm.Add_Resize({
        # 1. Verificamos se você está fisicamente interagindo com o Preview
        $isUserInteracting = ([System.Windows.Forms.Form]::ActiveForm -eq $previewForm)
        
        if ($previewForm.WindowState -eq 'Maximized') {
            if ($previewMem.Active -eq "A") { $previewMem.A_Max = $true } else { $previewMem.B_Max = $true }
        } 
        elseif ($previewForm.WindowState -eq 'Normal' -and $isUserInteracting) {
            # O SEGREDO ABSOLUTO: A memória só é apagada se você interagir com a janela.
            if ($previewMem.Active -eq "A") { $previewMem.A_Max = $false } else { $previewMem.B_Max = $false }
        }
    }.GetNewClosure())

    $form.Add_Resize({
        # Quando o Clone Commander "acorda" da barra de tarefas...
        if ($form.WindowState -ne 'Minimized') {
            $shouldBeMax = if ($previewMem.Active -eq "A") { $previewMem.A_Max } else { $previewMem.B_Max }
            
            if ($shouldBeMax) {
                # Coloca a ordem de Maximizar para o FIM da fila do Windows
                $form.BeginInvoke([System.Action]{
                    if ($previewForm.Visible) { $previewForm.WindowState = 'Maximized' }
                }) | Out-Null
            }
        }
    }.GetNewClosure())
    # ====================================================================

    # 1. Componente de Imagem (MOTOR WPF - ACELERADO POR HARDWARE / GPU)
    Add-Type -AssemblyName WindowsFormsIntegration
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    $wpfImage = New-Object System.Windows.Controls.Image
    $wpfImage.Stretch = [System.Windows.Media.Stretch]::Uniform
    
    $pbPreview = New-Object System.Windows.Forms.Integration.ElementHost
    $pbPreview.Dock = "Fill"
    $pbPreview.Child = $wpfImage
    $pbPreview.Visible = $false
    $previewForm.Controls.Add($pbPreview)

    # ====================================================================
    # 1.5 NOVO COMPONENTE: VISUALIZADOR DE GIF (WINDOWS FORMS NATIVO)
    # ====================================================================
    $wfGifPreview = New-Object System.Windows.Forms.PictureBox
    $wfGifPreview.Dock = "Fill"
    $wfGifPreview.SizeMode = "Zoom"
    $wfGifPreview.Visible = $false
    $previewForm.Controls.Add($wfGifPreview)
    # ====================================================================

    # 2. Componente de Texto/Codigo
    $rtbPreview = New-Object System.Windows.Forms.RichTextBox
    $rtbPreview.Dock = "Fill"
    $rtbPreview.ReadOnly = $true
    $rtbPreview.BackColor = "#1E1E1E"
    $rtbPreview.ForeColor = "#D4D4D4"
    $rtbPreview.Font = New-Object System.Drawing.Font("Consolas", 10)
    $rtbPreview.Visible = $false
    $previewForm.Controls.Add($rtbPreview)

    # 3. Componente de Audio e Video (WMP Otimizado)
    $wmpPreview = New-Object WMPHost
    $wmpPreview.Dock = "Fill"
    $wmpPreview.Visible = $false
    $previewForm.Controls.Add($wmpPreview)

    # ====================================================================
    # 4. COMPONENTE WEBVIEW2 (O MOTOR DO EDGE SOB DEMANDA)
    # ====================================================================
    $baseDir = $global:AppRoot
    $configDir = Join-Path $baseDir "config\WebView2"
    $wv2Dll = Join-Path $configDir "Microsoft.Web.WebView2.WinForms.dll"
    $wv2CoreDll = Join-Path $configDir "Microsoft.Web.WebView2.Core.dll"

    $wv2Preview = $null
    if ((Test-Path $wv2Dll) -and (Test-Path $wv2CoreDll)) {
        try {
            if ($env:Path -notmatch [regex]::Escape($configDir)) { $env:Path += ";$configDir" }
            $env:WEBVIEW2_USER_DATA_FOLDER = Join-Path $configDir "Profile"
            
            Add-Type -Path $wv2CoreDll -ErrorAction SilentlyContinue
            Add-Type -Path $wv2Dll -ErrorAction SilentlyContinue
            
            $wv2Preview = New-Object Microsoft.Web.WebView2.WinForms.WebView2
            $wv2Preview.Dock = "Fill"
            $wv2Preview.Visible = $false
            
            $wv2Preview.add_NavigationCompleted({
                param($sender, $e)
                try {
                    $js = "var v = document.querySelector('video'); if(v) { v.preload = 'auto'; v.loop = true; v.pause(); v.oncanplaythrough = function() { v.play(); }; }"
                    [void]$sender.ExecuteScriptAsync($js)
                } catch {}
            }.GetNewClosure())

            $previewForm.Controls.Add($wv2Preview)
        } catch { }
    }
    # ====================================================================

    # 5. Componente de Mensagem (Padrao)
    $lblPreviewMsg = New-Object System.Windows.Forms.Label
    $lblPreviewMsg.Dock = "Fill"
    $lblPreviewMsg.TextAlign = "MiddleCenter"
    $lblPreviewMsg.Font = New-Object System.Drawing.Font("Segoe UI", 12)
    $lblPreviewMsg.Text = "Selecione um arquivo..."
    $previewForm.Controls.Add($lblPreviewMsg)

    # 6. NOVO: BARRA DE STATUS DO PREVIEW (Resolução e Tamanho)
    $lblPreviewStatus = New-Object System.Windows.Forms.Label
    $lblPreviewStatus.Dock = "Bottom"
    $lblPreviewStatus.Height = 28
    $lblPreviewStatus.BackColor = "#2D2D30" # Cinza escuro elegante
    $lblPreviewStatus.ForeColor = "#E0E0E0" # Texto claro
    $lblPreviewStatus.TextAlign = "MiddleCenter"
    $lblPreviewStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblPreviewStatus.Text = ""
    
    # 1. AGORA FICA SEMPRE VISÍVEL
    $lblPreviewStatus.Visible = $true 
    
    $previewForm.Controls.Add($lblPreviewStatus)
    
    # 2. BLINDAGEM DE LAYOUT: Garante que nenhuma imagem ou texto cubra a barra
    $lblPreviewStatus.BringToFront()

    # TRAVA ANTI-FANTASMA: Limpeza total
    $previewForm.Add_FormClosing({
        param($s, $e)
        if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
            $e.Cancel = $true
            
            # Desliga o WMP em segurança
            try {
                $ocx = $wmpPreview.GetMediaPlayer()
                $ocx.controls.stop()
                $ocx.URL = ""
            } catch {}
            
            # Desliga o WebView2 (Edge) em segurança de forma isolada
            try {
                if ($wv2Preview -ne $null) { $wv2Preview.Source = New-Object System.Uri("about:blank") }
            } catch {}
            
            $s.Hide()
        }
    }.GetNewClosure())

    # ====================================================================
    # CORREÇÃO DE LEAK: DESTRUIÇÃO DEFINITIVA DOS PREVIEWS (MATA OS ZUMBIS)
    # Impede que o msedge.exe e o WMP fiquem rodando ocultos na memória
    # ====================================================================
    $form.Add_FormClosing({
        try {
            if ($wv2Preview -ne $null) { $wv2Preview.Dispose() }
            if ($wmpPreview -ne $null) { $wmpPreview.Dispose() }
            if ($previewForm -ne $null) { $previewForm.Dispose() }
        } catch {}
    }.GetNewClosure())

    $btnPreview.Add_Click({
        if ($previewForm.Visible) {
            # Desliga o WMP isolado
            try {
                $ocx = $wmpPreview.GetMediaPlayer()
                $ocx.controls.stop()
                $ocx.URL = ""
            } catch {}
            
            # Desliga o WebView2 isolado
            try {
                if ($wv2Preview -ne $null) { $wv2Preview.Source = New-Object System.Uri("about:blank") }
            } catch {}
            
            $previewForm.Hide()
        } else {
            $spawnPoint = $panel.PointToScreen([System.Drawing.Point]::new([int]($panel.Width / 4), 50))
            $previewForm.Location = $spawnPoint
            $parentForm = $panel.FindForm()
            if ($parentForm) { $previewForm.Show($parentForm) } else { $previewForm.Show() }
        }
        
        # TRUQUE NINJA: Limpa o destaque do botao de Preview
        Restore-ExplorerFocus
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

    # ====================================================================
    # --- BOTÃO RECORTAR ---
    # ====================================================================
    $btnCut = New-Object System.Windows.Forms.Button
    $btnCut.Dock = "Left"
    $btnCut.Width = 26
    $btnCut.FlatStyle = "Popup" 
    $btnCut.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCut.Text = ""
    $bookmarksBar.Controls.Add($btnCut)
    $btnCut.BringToFront()

    $ttCut = New-Object System.Windows.Forms.ToolTip
    $ttCut.SetToolTip($btnCut, "Recortar (Ctrl+X)")

    $dllShell32 = "$env:windir\System32\shell32.dll"
    try {
        $hIconCut = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllShell32, 259)
        if ($hIconCut -ne [IntPtr]::Zero) {
            $icoCut = [System.Drawing.Icon]::FromHandle($hIconCut)
            $origBmpCut = $icoCut.ToBitmap()
            $newSizeCut = New-Object System.Drawing.Size(16, 16)
            $btnCut.Image = New-Object System.Drawing.Bitmap($origBmpCut, $newSizeCut)
            
            $origBmpCut.Dispose()
            $icoCut.Dispose()
            [IconExtractor]::DestroyIcon($hIconCut) | Out-Null
        } else { $btnCut.Text = "X" }
    } catch { $btnCut.Text = "X" }

    $btnCut.Add_Click({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            try {
                $b = $tabControl.SelectedTab.Controls[0]
                $view = $b.ActiveXInstance.Document
                $items = $view.SelectedItems()
                if ($items.Count -gt 0) {
                    $files = New-Object System.Collections.Specialized.StringCollection
                    foreach ($item in $items) { $files.Add($item.Path) }
                    
                    $dataObj = New-Object System.Windows.Forms.DataObject
                    $dataObj.SetFileDropList($files)
                    
                    $dropEffect = New-Object byte[] 4
                    $dropEffect[0] = 2 
                    $memStream = New-Object System.IO.MemoryStream
                    $memStream.Write($dropEffect, 0, 4)
                    $dataObj.SetData("Preferred DropEffect", $memStream)
                    
                    [System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)
                }
                # --- FAXINA DE MEMÓRIA ---
                if ($items) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null }
                if ($view) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null }
            } catch {}
        }
        Restore-ExplorerFocus
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer


# ====================================================================
# --- BOTÃO COPIAR ---
# ====================================================================
    $btnCopyFile = New-Object System.Windows.Forms.Button
    $btnCopyFile.Dock = "Left"
    $btnCopyFile.Width = 26
    $btnCopyFile.FlatStyle = "Popup" 
    $btnCopyFile.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCopyFile.Text = ""
    $bookmarksBar.Controls.Add($btnCopyFile)
    $btnCopyFile.BringToFront()

    $ttCopyFile = New-Object System.Windows.Forms.ToolTip
    $ttCopyFile.SetToolTip($btnCopyFile, "Copiar (Ctrl+C)")

    try {
        $hIconCopy = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllShell32, 134)
        if ($hIconCopy -ne [IntPtr]::Zero) {
            $icoCopy = [System.Drawing.Icon]::FromHandle($hIconCopy)
            $origBmpCopy = $icoCopy.ToBitmap()
            $newSizeCopy = New-Object System.Drawing.Size(16, 16)
            $btnCopyFile.Image = New-Object System.Drawing.Bitmap($origBmpCopy, $newSizeCopy)
            
            $origBmpCopy.Dispose()
            $icoCopy.Dispose()
            [IconExtractor]::DestroyIcon($hIconCopy) | Out-Null
        } else { $btnCopyFile.Text = "C" }
    } catch { $btnCopyFile.Text = "C" }

    $btnCopyFile.Add_Click({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            try {
                $b = $tabControl.SelectedTab.Controls[0]
                $view = $b.ActiveXInstance.Document
                $items = $view.SelectedItems()
                if ($items.Count -gt 0) {
                    $files = New-Object System.Collections.Specialized.StringCollection
                    foreach ($item in $items) { $files.Add($item.Path) }
                    
                    $dataObj = New-Object System.Windows.Forms.DataObject
                    $dataObj.SetFileDropList($files)
                    
                    $dropEffect = New-Object byte[] 4
                    $dropEffect[0] = 5 
                    $memStream = New-Object System.IO.MemoryStream
                    $memStream.Write($dropEffect, 0, 4)
                    $dataObj.SetData("Preferred DropEffect", $memStream)
                    
                    [System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)
                }
                # --- FAXINA DE MEMÓRIA ---
                if ($items) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null }
                if ($view) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null }
            } catch {}
        }
        Restore-ExplorerFocus
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

# ====================================================================
# --- BOTÃO COLAR (USANDO VERBO NATIVO) ---
# ====================================================================
    $btnPaste = New-Object System.Windows.Forms.Button
    $btnPaste.Dock = "Left"
    $btnPaste.Width = 26
    $btnPaste.FlatStyle = "Popup" 
    $btnPaste.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnPaste.Text = ""
    $bookmarksBar.Controls.Add($btnPaste)
    $btnPaste.BringToFront()

    $ttPaste = New-Object System.Windows.Forms.ToolTip
    $ttPaste.SetToolTip($btnPaste, "Colar (Ctrl+V)")

    try {
        $hIconPaste = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllShell32, 260)
        if ($hIconPaste -ne [IntPtr]::Zero) {
            $icoPaste = [System.Drawing.Icon]::FromHandle($hIconPaste)
            $origBmpPaste = $icoPaste.ToBitmap()
            $newSizePaste = New-Object System.Drawing.Size(16, 16)
            $btnPaste.Image = New-Object System.Drawing.Bitmap($origBmpPaste, $newSizePaste)
            
            $origBmpPaste.Dispose()
            $icoPaste.Dispose()
            [IconExtractor]::DestroyIcon($hIconPaste) | Out-Null
        } else { $btnPaste.Text = "V" }
    } catch { $btnPaste.Text = "V" }

    $btnPaste.Add_Click({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            try {
                $b = $tabControl.SelectedTab.Controls[0]
                $view = $b.ActiveXInstance.Document
                
                # Desmarca os arquivos selecionados para não colar acidentalmente dentro deles
                $items = $view.SelectedItems()
                if ($items.Count -gt 0) {
                    foreach ($item in $items) { $view.SelectItem($item, 0) }
                }
                
                # A MÁGICA: Manda a pasta executar o "Colar" nativo dela mesma
                $folder = $view.Folder
                $self = $folder.Self
                $self.InvokeVerb("paste")
                
                # --- FAXINA DE MEMÓRIA ---
                if ($items) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null }
                if ($self) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($self) | Out-Null }
                if ($folder) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null }
                if ($view) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null }
            } catch {}
        }
        Restore-ExplorerFocus
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

# --- BOTÃO EXCLUIR (DEFINITIVO COM ÍCONE NETSHELL) ---
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Dock = "Left"
    $btnDelete.Width = 26
    
    # ESTILO ATUALIZADO PARA POPUP
    $btnDelete.FlatStyle = "Popup" 
    
    $btnDelete.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnDelete.Text = ""
    $bookmarksBar.Controls.Add($btnDelete)
    $btnDelete.BringToFront()

    $ttDelete = New-Object System.Windows.Forms.ToolTip
    $ttDelete.SetToolTip($btnDelete, "Excluir selecionados (Mover para a Lixeira)")

    # Extração do ícone escolhido: netshell.dll, índice 25
    $dllNetShell = "$env:windir\System32\netshell.dll"
    try {
        $hIconDel = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllNetShell, 25)
        if ($hIconDel -ne [IntPtr]::Zero) {
            $icoDel = [System.Drawing.Icon]::FromHandle($hIconDel)
            $origBmpDel = $icoDel.ToBitmap()
            
            $newSizeDel = New-Object System.Drawing.Size(16, 16)
            $btnDelete.Image = New-Object System.Drawing.Bitmap($origBmpDel, $newSizeDel)
            
            # --- LIMPEZA DE MEMÓRIA (Vazamento corrigido) ---
            $origBmpDel.Dispose()
            $icoDel.Dispose()
            [IconExtractor]::DestroyIcon($hIconDel) | Out-Null
        } else { 
            $btnDelete.Text = "X" 
        }
    } catch { 
        $btnDelete.Text = "X" 
    }

    # Ação de Excluir Arquivos (Motor nativo do Windows)
    $btnDelete.Add_Click({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]
            try {
                $view = $b.ActiveXInstance.Document
                $items = $view.SelectedItems()
                
                if ($items.Count -gt 0) {
                    # Manda o Windows executar a ação nativa de exclusão
                    foreach ($item in $items) {
                        $item.InvokeVerb("delete")
                    }
                }
                # --- FAXINA DE MEMÓRIA ---
                if ($items) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null }
                if ($view) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null }
            } catch {}
        }
        
        # TRUQUE NINJA: Tira o foco do botão após clicar em excluir
        Restore-ExplorerFocus
    }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

    # ====================================================================
    # --- BOTÃO BACKUP (NOVO - MOTOR ASSÍNCRONO COM RUNSPACE POOL) ---
    # ====================================================================
    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Dock = "Left"
    $btnBackup.Width = 26  # Quadrado perfeito
    $btnBackup.FlatStyle = "Popup" 
    $btnBackup.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBackup.Text = "" 
    $bookmarksBar.Controls.Add($btnBackup)
    $btnBackup.BringToFront()

    $ttBackup = New-Object System.Windows.Forms.ToolTip
    $ttBackup.SetToolTip($btnBackup, "Backup Inteligente`nCria ou atualiza versão (1), (2)... se houver mudanças.")

    $dllPif = "$env:windir\System32\pifmgr.dll"
    try {
        $hIconBkp = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllPif, 13)
        if ($hIconBkp -ne [IntPtr]::Zero) {
            $icoBkp = [System.Drawing.Icon]::FromHandle($hIconBkp)
            $origBmp = $icoBkp.ToBitmap()
            $newSize = New-Object System.Drawing.Size(16, 16)
            $btnBackup.Image = New-Object System.Drawing.Bitmap($origBmp, $newSize)
            $origBmp.Dispose(); $icoBkp.Dispose()
            [IconExtractor]::DestroyIcon($hIconBkp) | Out-Null
        } else { $btnBackup.Text = "B" }
    } catch { $btnBackup.Text = "B" }

    $btnBackup.Add_Click({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]
            try {
                $view = $b.ActiveXInstance.Document
                $selectedItems = $view.SelectedItems()
                
                if ($selectedItems.Count -eq 0) { 
                    Restore-ExplorerFocus 
                    if ($selectedItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selectedItems) | Out-Null }
                    if ($view) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null }
                    return 
                }

                # 1. EMPACOTA OS DADOS PARA ENVIAR PARA A THREAD (Evita crash de COM em thread secundária)
                $itemsToProcess = @()
                foreach ($item in $selectedItems) {
                    $itemsToProcess += @{ Path = $item.Path; IsFolder = $item.IsFolder; Name = $item.Name }
                }
                
                if ($selectedItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selectedItems) | Out-Null }
                if ($view) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null }

                # 2. FEEDBACK VISUAL
                $btnBackup.BackColor = "#FF8C00" # Laranja indicando trabalho no fundo
                $btnBackup.Enabled = $false
                
                # 3. MISSÃO PARA O ESQUADRÃO TÁTICO
                $bgScript = {
                    param($items, $form, $btn)
                    
                    foreach ($item in $items) {
                        $srcPath = $item.Path
                        $isFolder = $item.IsFolder
                        $itemName = $item.Name
                        
                        $parentPath = Split-Path $srcPath -Parent
                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($srcPath)
                        $ext = [System.IO.Path]::GetExtension($srcPath)
                        if ($isFolder) { $baseName = $itemName; $ext = "" }

                        $backupDir = Join-Path $parentPath "$itemName - Backup"
                        if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }

                        function Get-Signature($path, $isDir) {
                            if (-not (Test-Path $path)) { return @{ Size = -1; Date = [DateTime]::MinValue } }
                            if ($isDir) {
                                $files = Get-ChildItem $path -Recurse -File -Force -ErrorAction SilentlyContinue
                                $size = if ($files) { ($files | Measure-Object -Property Length -Sum).Sum } else { 0 }
                                $date = if ($files) { ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime } else { (Get-Item $path).LastWriteTime }
                                return @{ Size = $size; Date = $date }
                            } else {
                                $file = Get-Item $path
                                return @{ Size = $file.Length; Date = $file.LastWriteTime }
                            }
                        }

                        $srcSig = Get-Signature $srcPath $isFolder
                        $lastBackupIndex = -1
                        $lastBackupPath = ""
                        
                        $existingBackups = Get-ChildItem $backupDir | Where-Object { $_.Name -match "^$([regex]::Escape($baseName))( \(\d+\))?$([regex]::Escape($ext))$" }
                        
                        if ($existingBackups) {
                            foreach ($eb in $existingBackups) {
                                if ($eb.Name -eq $itemName) { 
                                    if ($lastBackupIndex -lt 0) { $lastBackupIndex = 0; $lastBackupPath = $eb.FullName }
                                } elseif ($eb.Name -match "\((\d+)\)$([regex]::Escape($ext))$") {
                                    $num = [int]$matches[1]
                                    if ($num -gt $lastBackupIndex) {
                                        $lastBackupIndex = $num
                                        $lastBackupPath = $eb.FullName
                                    }
                                }
                            }

                            if ($lastBackupIndex -ge 0 -and (Test-Path $lastBackupPath)) {
                                $tgtSig = Get-Signature $lastBackupPath $isFolder
                                $timeDiff = [math]::Abs(($srcSig.Date - $tgtSig.Date).TotalSeconds)
                                if ($srcSig.Size -eq $tgtSig.Size -and $timeDiff -lt 2) { continue }
                            }
                        }

                        $lastBackupIndex++
                        $newBackupName = if ($lastBackupIndex -eq 0) { $itemName } else { "$baseName ($lastBackupIndex)$ext" }
                        $destPath = Join-Path $backupDir $newBackupName

                        if ($isFolder) { Copy-Item -Path $srcPath -Destination $destPath -Recurse -Force } 
                        else { Copy-Item -Path $srcPath -Destination $destPath -Force }
                    }
                    
                    # 4. AVISA O GERENTE QUE TERMINOU
                    $form.Invoke([System.Action]{
                        $btn.BackColor = "Transparent"
                        $btn.Enabled = $true
                        [System.Windows.Forms.MessageBox]::Show("Backup Inteligente concluído em segundo plano!", "Tarefa Concluida", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    })
                }
                
                $ps = [powershell]::Create()
                $ps.RunspacePool = $global:RunspacePool
                [void]$ps.AddScript($bgScript)
                [void]$ps.AddArgument($itemsToProcess)
                [void]$ps.AddArgument($global:SyncHash.Form)
                [void]$ps.AddArgument($btnBackup)
                [void]$ps.BeginInvoke()
                
            } catch {}
        }
        Restore-ExplorerFocus
    }.GetNewClosure())
    
    # ====================================================================
    # --- MENU SUSPENSO (Para o botão Novo) ---
    # ====================================================================
    $menuNew = New-Object System.Windows.Forms.ContextMenuStrip

    # Função auxiliar TOTALMENTE BLINDADA (Agora com Auto-Seleção e Renomear)
    $ActionCreateNew = {
        param($TControl, $Prefix, $Extension, $IsFolder)
        
        try {
            if ($TControl.SelectedTab -and $TControl.SelectedTab.Controls.Count -gt 0) {
                $activeBrowser = $TControl.SelectedTab.Controls[0]
                $currentPath = $activeBrowser.Url.LocalPath
                
                # Bloqueio de pastas virtuais
                if ([string]::IsNullOrWhiteSpace($currentPath) -or $currentPath -match "^::" -or $currentPath -match "^search-ms:") {
                    [System.Windows.Forms.MessageBox]::Show("Navegue até uma pasta real para criar arquivos.", "Aviso", 0, 48)
                    return
                }

                # Lógica de nomes (Novo Arquivo, Novo Arquivo (1)...)
                $baseName = "$Prefix$Extension"
                $fullPath = Join-Path $currentPath $baseName
                $count = 1
                while (Test-Path $fullPath) {
                    $baseName = "$Prefix ($count)$Extension"
                    $fullPath = Join-Path $currentPath $baseName
                    $count++
                }

                # Criação com trava
                if ($IsFolder) {
                    New-Item -ItemType Directory -Path $fullPath -ErrorAction Stop | Out-Null
                } else {
                    New-Item -ItemType File -Path $fullPath -ErrorAction Stop | Out-Null
                }
                
                # ====================================================================
                # NOVA MÁGICA: AUTO-SELEÇÃO E MODO RENOMEAR (Sem dar Refresh!)
                # ====================================================================
                $view = $activeBrowser.ActiveXInstance.Document
                if ($view) {
                    $folder = $view.Folder
                    if ($folder) {
                        $fileName = Split-Path $fullPath -Leaf
                        
                        $retries = 0
                        $item = $null
                        while ($null -eq $item -and $retries -lt 10) {
                            $item = $folder.ParseName($fileName)
                            if ($null -eq $item) { Start-Sleep -Milliseconds 50; $retries++ }
                        }

                        if ($item) {
                            $view.SelectItem($item, 15)
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null 
                        }
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                    }
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
                }
                
            } else {
                [System.Windows.Forms.MessageBox]::Show("Nenhuma aba ativa encontrada.", "Aviso", 0, 48)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro:`n$_", "Erro do Sistema", 0, 16)
        }
    }

    # Adicionando opções e injetando a variável $tabControl explicitamente para não dar "amnésia"
    $menuNew.Items.Add("Nova Pasta").Add_Click({ & $ActionCreateNew -TControl $tabControl -Prefix "Nova Pasta" -Extension "" -IsFolder $true }.GetNewClosure())
    [void]$menuNew.Items.Add("-") # Linha divisória
    $menuNew.Items.Add("Novo Texto (.txt)").Add_Click({ & $ActionCreateNew -TControl $tabControl -Prefix "Novo Arquivo" -Extension ".txt" -IsFolder $false }.GetNewClosure())
    $menuNew.Items.Add("Novo Script (.ps1)").Add_Click({ & $ActionCreateNew -TControl $tabControl -Prefix "Novo Script" -Extension ".ps1" -IsFolder $false }.GetNewClosure())
    $menuNew.Items.Add("Nova Planilha (.csv)").Add_Click({ & $ActionCreateNew -TControl $tabControl -Prefix "Nova Planilha" -Extension ".csv" -IsFolder $false }.GetNewClosure())
    $menuNew.Items.Add("Novo Batch (.bat)").Add_Click({ & $ActionCreateNew -TControl $tabControl -Prefix "Novo Batch" -Extension ".bat" -IsFolder $false }.GetNewClosure())

    # --- ESPAÇADOR INVISÍVEL ---
    Add-Spacer

    # ====================================================================
    # --- BOTÃO NOVO (Com sinal de + em negrito) ---
    # ====================================================================
    $btnNew = New-Object System.Windows.Forms.Button
    $btnNew.Dock = "Left"
    $btnNew.Width = 26  # Quadrado perfeito, igual ao de Backup
    
    # ESTILO ATUALIZADO PARA POPUP
    $btnNew.FlatStyle = "Popup" 
    
    $btnNew.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # NOVA LÓGICA: Usando o sinal de + e definindo a fonte como Negrito
    $btnNew.Text = "+"
    $btnNew.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    
    $bookmarksBar.Controls.Add($btnNew)
    $btnNew.BringToFront()

    $ttNew = New-Object System.Windows.Forms.ToolTip
    $ttNew.SetToolTip($btnNew, "Novo Item (Pasta, Txt, Script, CSV...)")

    # Em vez de criar algo direto, o clique abre o Menu Suspenso debaixo do botão
    $btnNew.Add_Click({
        $menuNew.Show($btnNew, 0, $btnNew.Height)
        
        # TRUQUE NINJA: Tira o foco do botão após abrir o menu
        Restore-ExplorerFocus
    }.GetNewClosure())

    # --- ESPAÇADOR 1: ENTRE MARCADORES E NAVEGAÇÃO (2 PIXELS) ---
    $spacer1 = New-Object System.Windows.Forms.Panel
    $spacer1.Dock = "Top"; $spacer1.Height = 2
    $topContainer.Controls.Add($spacer1)

    # ====================================================================
    # 2.2 NAV BAR (RENDERIZADA ABAIXO DOS MARCADORES)
    # ====================================================================
    $navBar = New-Object System.Windows.Forms.Panel; $navBar.Dock = "Top"; $navBar.Height = 26; $topContainer.Controls.Add($navBar)
    
    # --- BOTAO DE FAVORITO (Na Direita) ---
    $btnFav = New-Object System.Windows.Forms.Button
    $btnFav.Dock = "Right"
    $btnFav.Width = 35
    $btnFav.FlatStyle = "Flat"
    $btnFav.FlatAppearance.BorderColor = "Black"
    $btnFav.FlatAppearance.BorderSize = 1
    $btnFav.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnFav.Text = ""
    if ($global:StarEmptyIconBmp) { $btnFav.Image = $global:StarEmptyIconBmp } else { $btnFav.Text = "F" }
    $navBar.Controls.Add($btnFav)
    $btnFav.BringToFront()

    # --- NOVO: ESPACADOR DO FAVORITO (1 Pixel na Direita) ---
    $navSpFav = New-Object System.Windows.Forms.Label
    $navSpFav.AutoSize = $false
    $navSpFav.Width = 1
    $navSpFav.Dock = "Right"
    $navSpFav.BackColor = [System.Drawing.Color]::Transparent
    $navBar.Controls.Add($navSpFav)
    $navSpFav.BringToFront()

    # --- BOTOES DE NAVEGACAO (Na Esquerda) ---
    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Dock = "Left"
    $btnBack.Width = 35
    $btnBack.FlatStyle = "Flat"
    $btnBack.FlatAppearance.BorderColor = "Black"
    $btnBack.FlatAppearance.BorderSize = 1
    $btnBack.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBack.Text = "<"
    $btnBack.Enabled = $false
    $navBar.Controls.Add($btnBack)
    $btnBack.BringToFront()

    $navSp1 = New-Object System.Windows.Forms.Label; $navSp1.AutoSize = $false; $navSp1.Width = 1; $navSp1.Dock = "Left"; $navSp1.BackColor = [System.Drawing.Color]::Transparent; $navBar.Controls.Add($navSp1); $navSp1.BringToFront()

    $btnFwd = New-Object System.Windows.Forms.Button
    $btnFwd.Dock = "Left"
    $btnFwd.Width = 35
    $btnFwd.FlatStyle = "Flat"
    $btnFwd.FlatAppearance.BorderColor = "Black"
    $btnFwd.FlatAppearance.BorderSize = 1
    $btnFwd.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnFwd.Text = ">"
    $btnFwd.Enabled = $false
    $navBar.Controls.Add($btnFwd)
    $btnFwd.BringToFront()

    $navSp2 = New-Object System.Windows.Forms.Label; $navSp2.AutoSize = $false; $navSp2.Width = 1; $navSp2.Dock = "Left"; $navSp2.BackColor = [System.Drawing.Color]::Transparent; $navBar.Controls.Add($navSp2); $navSp2.BringToFront()

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Dock = "Left"
    $btnUp.Width = 35
    $btnUp.FlatStyle = "Flat"
    $btnUp.FlatAppearance.BorderColor = "Black"
    $btnUp.FlatAppearance.BorderSize = 1
    $btnUp.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnUp.Text = "^"
    $btnUp.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
    $btnUp.Enabled = $true
    $navBar.Controls.Add($btnUp)
    $btnUp.BringToFront()

    # --- CORRECAO: ESPACADOR DA BARRA DE ENDERECO (1 Pixel na Esquerda) ---
    $navSp3 = New-Object System.Windows.Forms.Label; $navSp3.AutoSize = $false; $navSp3.Width = 1; $navSp3.Dock = "Left"; $navSp3.BackColor = [System.Drawing.Color]::Transparent; $navBar.Controls.Add($navSp3); $navSp3.BringToFront()

    # --- BARRA DE ENDERECO (CAIXA FALSA 100% BLINDADA) ---
    $pathWrapper = New-Object System.Windows.Forms.Panel
    $pathWrapper.Dock = "Fill"
    $pathWrapper.BackColor = "White"
    $pathWrapper.BorderStyle = "FixedSingle"
    $pathWrapper.Cursor = [System.Windows.Forms.Cursors]::IBeam 
    
    # FECHA A BRECHA 1: Oculta o painel do radar da tecla TAB
    $pathWrapper.TabStop = $false 
    
    $navBar.Controls.Add($pathWrapper)
    $pathWrapper.BringToFront()

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.BorderStyle = "None"
    $txtPath.Text = $InitialPath
    $txtPath.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtPath.Location = New-Object System.Drawing.Point(4, 3)
    $txtPath.Width = 200 

    $pathWrapper.Add_Resize({
        if ($txtPath -ne $null -and $pathWrapper -ne $null) {
            # FECHA A BRECHA 2: Impede que a largura fique negativa e cause Crash
            $newWidth = $pathWrapper.Width - 8
            if ($newWidth -gt 0) {
                $txtPath.Width = $newWidth
            }
        }
    }.GetNewClosure())

    # FECHA A BRECHA 3: Captura qualquer botao do mouse (Esquerdo/Direito) e foca no texto
    $pathWrapper.Add_MouseDown({
        param($sender, $e)
        $txtPath.Focus()
    }.GetNewClosure())

    $pathWrapper.Controls.Add($txtPath)

    # --- ESPAÇADOR 2: ENTRE NAVEGAÇÃO E ABAS (2 PIXELS) ---
    $spacer2 = New-Object System.Windows.Forms.Panel
    $spacer2.Dock = "Top"; $spacer2.Height = 2
    $topContainer.Controls.Add($spacer2)

    # --- LÓGICA DE CRIAÇÃO DE ABAS ---
    $AddTabLogic = {
        param($TargetTabControl, $Path, $AddressBox, $bBack, $bFwd, $BtnView)
        
        $targetPath = $Path
        $targetView = $null
        
        if ($Path -ne $null) {
            try {
                if ($Path.Path -ne $null) { $targetPath = $Path.Path }
                if ($Path.View -ne $null) { $targetView = $Path.View }
            } catch {}
        }
        
        # ====================================================================
        # 1. PREPARAÇÃO DO CAMINHO (Com a sua Proteção de Regressão)
        # ====================================================================
        [string]$safePath = $targetPath
        
        if ([string]::IsNullOrWhiteSpace($safePath) -or $safePath -match "System\." -or $safePath -match "@{") { 
            $safePath = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 
        }
        if ($safePath.StartsWith("::")) { $safePath = "shell:" + $safePath }

        # --- A SUA LÓGICA DE REGRESSÃO (FALLBACK MELHORADO) ---
        $isVirtual = ($safePath -match "^(shell|::|search\-ms)")
        if (-not $isVirtual -and -not [string]::IsNullOrWhiteSpace($safePath)) {
            
            # NOVO: Pergunta ao Hardware se o disco físico está realmente lá (Bloqueia leitores vazios)
            $isDriveReady = $true
            if ($safePath -match "^([A-Za-z]:)") {
                try { $isDriveReady = (New-Object System.IO.DriveInfo($matches[1])).IsReady } catch { $isDriveReady = $false }
            }

            if (-not $isDriveReady) {
                $safePath = "" # Mata o caminho imediatamente se não houver pendrive físico
            } else {
                # Fica cortando a última pasta até achar um caminho que o PC reconheça
                while (-not [string]::IsNullOrWhiteSpace($safePath) -and -not (Test-Path -LiteralPath $safePath -ErrorAction SilentlyContinue)) {
                    try { 
                        $parent = Split-Path $safePath -Parent -ErrorAction Stop
                        if ($parent -eq $safePath) { $safePath = ""; break } # Proteção contra loop
                        $safePath = $parent
                    } catch { $safePath = "" }
                }
            }
            
            # Se sumiu tudo (Pendrive removido), vai pro Meu Computador
            if ([string]::IsNullOrWhiteSpace($safePath)) {
                $safePath = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
            }
        }

        # Descobre o nome provisório inteligente antes mesmo da aba piscar na tela
        $tabName = "Carregando..."
        if ($safePath -match "20D04FE0-3AEA-1069-A2D8-08002B30309D") { $tabName = "Meu Computador" }
        elseif ($safePath -match "645FF040-5081-101B-9F08-00AA002F954E") { $tabName = "Lixeira" }
        elseif ($safePath -notmatch "^(shell|::)") {
            try { 
                $tabName = Split-Path $safePath -Leaf
                if ([string]::IsNullOrWhiteSpace($tabName)) { $tabName = $safePath } 
            } catch {}
        }
        if ($tabName.Length -gt 21) { $tabName = $tabName.Substring(0, 21) + "..." }

        # ====================================================================
        # 2. CRIACAO DA ABA (Ja com o nome correto)
        # ====================================================================
        $tab = New-Object System.Windows.Forms.TabPage; $tab.Text = "$tabName      "
        $browser = New-Object System.Windows.Forms.WebBrowser; $browser.Dock = "Fill"; $browser.ScriptErrorsSuppressed = $true
        
        $startView = "Detalhes"
        if ($targetView) {
            $startView = $targetView
        } elseif ($TargetTabControl.SelectedTab -ne $null -and $TargetTabControl.SelectedTab.Controls.Count -gt 0) {
            $activeB = $TargetTabControl.SelectedTab.Controls[0]
            if ($activeB.Tag -and $activeB.Tag.ViewMode) { $startView = $activeB.Tag.ViewMode }
        }
        
        $browser.Tag = New-Object PSObject -Property @{ ViewMode = $startView }
        
        $tab.Controls.Add($browser)
        $TargetTabControl.TabPages.Add($tab)
        $TargetTabControl.SelectedTab = $tab

        # --- NOVO GATILHO: MEMORIA INSTANTANEA DE SAIDA ---
        $browser.Add_Navigating({
            param($s, $e)
            try {
                $view = $s.ActiveXInstance.Document
                if ($view) {
                    $folder = $view.Folder
                    if ($folder) {
                        $selfObj = $folder.Self
                        if ($selfObj) {
                            $rawPath = $selfObj.Path
                            $items = $view.SelectedItems()
                            
                            if ($items.Count -gt 0) {
                                if ($null -eq $global:FolderSelMemory) { $global:FolderSelMemory = @{} }
                                $memArr = @()
                                foreach ($i in $items) { try { $memArr += $i.Name } catch {} }
                                $global:FolderSelMemory[$rawPath] = $memArr
                            }
                            
                            if ($items) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null }
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selfObj) | Out-Null
                        }
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                    }
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
                }
            } catch {}
        }.GetNewClosure())
        # --------------------------------------------------

        $browser.Add_Navigated({ param($s, $e)
            try {
                $p = $s.Url.LocalPath
                # So atualiza automaticamente aqui se NAO for uma pasta virtual
                if ($p -notmatch "^::" -and -not [string]::IsNullOrWhiteSpace($p)) {
                    $n = Split-Path $p -Leaf; if ([string]::IsNullOrEmpty($n)) { $n = $p }
                    if ($n.Length -gt 21) { $n = $n.Substring(0, 21) + "..." }
                    $tab.Text = "$n      " 
                }
                
                if ($TargetTabControl.SelectedTab -eq $tab) {
                    $AddressBox.Text = $p
                    $bBack.Enabled = $s.CanGoBack; $bFwd.Enabled = $s.CanGoForward
                    $global:ActiveBrowser = $s
                }
            } catch {}
        }.GetNewClosure())

        $browser.Add_DocumentCompleted({ 
            param($s, $e) 
            
            # --- GATILHO RECOVERY ---
            if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
        }.GetNewClosure())
        
        # ====================================================================
        # 3. NAVEGACAO FINAL
        # ====================================================================
        try {
            $browser.Navigate($safePath)
        } catch {
            $browser.Navigate("C:\")
        }
    }

    $tabControl.Tag = @{ 
        TxtPath = $txtPath; BtnBack = $btnBack; BtnFwd = $btnFwd
        AddTab = $AddTabLogic; BtnView = $btnViewMode
    }

    # ====================================================================
    # --- TRAVA DE SEGURANÇA (Bloqueia o botão "+" em pastas virtuais) ---
    # ====================================================================
    $CheckNewButtonState = {
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]
            $p = $b.Url.LocalPath
            
            # Se for pasta virtual (Lixeira, Meu Computador, etc)
            if ([string]::IsNullOrWhiteSpace($p) -or $p -match "^::" -or $p -match "^search-ms:") {
                $btnNew.Enabled = $false
            } else {
                $btnNew.Enabled = $true
            }
        }
    }

    # Gatilho 1: Aciona a verificação toda vez que a barra de endereços mudar (Navegação)
    $txtPath.Add_TextChanged({
        & $CheckNewButtonState
    }.GetNewClosure())

    # Gatilho 2: Aciona a verificação toda vez que você trocar de aba clicando nelas
    $tabControl.Add_SelectedIndexChanged({
        & $CheckNewButtonState
    }.GetNewClosure())
    # ====================================================================

    # ====================================================================
    # --- BOTÃO "+" FIXO E INDICADORES
    # ====================================================================
    $btnNewTab = New-Object System.Windows.Forms.Button
    $btnNewTab.Text = "+"
    $btnNewTab.Size = New-Object System.Drawing.Size(20, 20) 
    $btnNewTab.UseVisualStyleBackColor = $true
    $btnNewTab.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $btnNewTab.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnNewTab.Location = New-Object System.Drawing.Point(($panel.Width - 22), 56)
    
    $panel.Controls.Add($btnNewTab)
    $btnNewTab.BringToFront()

    $btnNewTab.Add_Click({ 
        & $AddTabLogic -TargetTabControl $tabControl -Path "C:\" -AddressBox $txtPath -bBack $btnBack -bFwd $btnFwd -BtnView $btnViewMode
        # --- GATILHO RECOVERY: Salva ao abrir nova aba no + ---
        if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
    }.GetNewClosure())

    # --- INDICADOR ESQUERDO DE ABAS OCULTAS ---
    $lblMoreLeft = New-Object System.Windows.Forms.Label
    $lblMoreLeft.Text = [char]0x00AB
    $lblMoreLeft.Size = New-Object System.Drawing.Size(20, 20)
    $lblMoreLeft.Location = New-Object System.Drawing.Point(0, 56)
    $lblMoreLeft.BackColor = "#1E1E1E" 
    $lblMoreLeft.ForeColor = "White"
    $lblMoreLeft.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
    $lblMoreLeft.TextAlign = "MiddleCenter"
    $lblMoreLeft.Visible = $false
    
    $ttLeft = New-Object System.Windows.Forms.ToolTip
    $ttLeft.SetToolTip($lblMoreLeft, "Existem abas ocultas a esquerda")
    
    $panel.Controls.Add($lblMoreLeft)
    $lblMoreLeft.BringToFront()

    # --- DESENHO DAS ABAS (COM QUADRADO NO X E CORREÇÃO DE MEMÓRIA) ---
    $tabControl.Add_DrawItem({ param($s, $e)
        try {
            if ($e.Index -lt 0 -or $e.Index -ge $s.TabPages.Count) { return }
            $g = $e.Graphics; $r = $s.GetTabRect($e.Index)
            if ($r.Width -le 0) { return }
            
            $rectF = [System.Drawing.RectangleF]::new([float]$r.X, [float]$r.Y, [float]$r.Width, [float]$r.Height)
            $isSelected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected)
            
            if ($isSelected) {
                $bgBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(80, 80, 85))
                $textBrush = [System.Drawing.Brushes]::White
                $xBrush = [System.Drawing.Brushes]::White
                $boxBgBrush = [System.Drawing.Brushes]::Red
                $borderPen = [System.Drawing.Pens]::DarkRed
                $fontStyle = [System.Drawing.FontStyle]::Bold
            } else {
                $bgBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(240, 240, 240))
                $textBrush = [System.Drawing.Brushes]::DimGray
                $xBrush = [System.Drawing.Brushes]::Gray
                $boxBgBrush = $null
                $borderPen = [System.Drawing.Pens]::LightGray
                $fontStyle = [System.Drawing.FontStyle]::Regular
            }

            $g.FillRectangle($bgBrush, $rectF)
            
            $font = [System.Drawing.Font]::new($s.Font, $fontStyle)
            
            $sfText = New-Object System.Drawing.StringFormat
            $sfText.Alignment = "Near"
            $sfText.LineAlignment = "Center"
            $sfText.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
            $sfText.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
            
            $title = $s.TabPages[$e.Index].Text
            $textRect = [System.Drawing.RectangleF]::new([float]($r.X + 5), [float]$r.Y, [float]($r.Width - 22), [float]$r.Height)
            $g.DrawString($title, $font, $textBrush, $textRect, $sfText)
                  
            $boxSize = 14.0
            $boxX = [float]($r.Right - 18)
            $boxY = [float]($r.Top + (($r.Height - $boxSize) / 2))
            $closeBoxRect = [System.Drawing.RectangleF]::new($boxX, $boxY, $boxSize, $boxSize)
            
            if ($boxBgBrush -ne $null) {
                $g.FillRectangle($boxBgBrush, $closeBoxRect.X, $closeBoxRect.Y, $closeBoxRect.Width, $closeBoxRect.Height)
            }
            $g.DrawRectangle($borderPen, $closeBoxRect.X, $closeBoxRect.Y, $closeBoxRect.Width, $closeBoxRect.Height)
            
            $xFont = [System.Drawing.Font]::new("Consolas", 8, [System.Drawing.FontStyle]::Bold)
            $sfX = New-Object System.Drawing.StringFormat; $sfX.Alignment = "Center"; $sfX.LineAlignment = "Center"
            $g.DrawString("x", $xFont, $xBrush, $closeBoxRect, $sfX)
            
            $bgBrush.Dispose()
            
            # ==============================================================
            # CORREÇÃO: Limpando os objetos GDI recém-criados
            # ==============================================================
            $font.Dispose()
            $sfText.Dispose()
            $xFont.Dispose()
            $sfX.Dispose()
            
        } catch {}
    })

    # --- FECHAR ABA (CLIQUE NO X) E PREPARAR ARRASTO ---
    $tabControl.Add_MouseDown({ param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $s.Tag.DragStart = $e.Location
        }
        
        for ($i = 0; $i -lt $s.TabPages.Count; $i++) {
            $r = $s.GetTabRect($i)
            
            $boxSize = 14
            $boxX = $r.Right - 18
            $boxY = $r.Top + [int](($r.Height - $boxSize) / 2)
            $closeRect = [System.Drawing.Rectangle]::new($boxX, $boxY, $boxSize, $boxSize)
            
            if ($closeRect.Contains($e.Location)) {
                $s.Tag.DragStart = $null 
                
                # ====================================================================
                # LIXEIRO SEGURO (Passo 1: Guarda quem vai morrer e tira da tela)
                # ====================================================================
                $tabToClose = $s.TabPages[$i]
                $currentIndex = $i
                
                # MÁGICA DE NAVEGADOR: Pula para a aba vizinha ANTES de deletar a atual
                if ($s.TabPages.Count -gt 1 -and $s.SelectedTab -eq $tabToClose) {
                    if ($currentIndex -eq ($s.TabPages.Count - 1)) {
                        $s.SelectedIndex = $currentIndex - 1 # Última aba -> vai pra esquerda
                    } else {
                        $s.SelectedIndex = $currentIndex + 1 # Tem aba à direita -> foca nela
                    }
                }

                $s.TabPages.Remove($tabToClose) # A aba vizinha assume com 100% de segurança aqui
                
                # Se era a última aba, recria o "Meu Computador"
                if ($s.TabPages.Count -eq 0) { 
                    & ($s.Tag.AddTab) -TargetTabControl $s -Path "C:\" -AddressBox ($s.Tag.TxtPath) -bBack ($s.Tag.BtnBack) -bFwd ($s.Tag.BtnFwd) -BtnView ($s.Tag.BtnView)
                }

                # (Passo 2 e 3: Destrói o navegador nos bastidores de forma limpa)
                try {
                    if ($tabToClose.Controls.Count -gt 0) {
                        $browserToKill = $tabToClose.Controls[0]
                        if ($global:ActiveBrowser -eq $browserToKill) { $global:ActiveBrowser = $null }
                        
                        # --- CORRECAO DE MEMORY LEAK (O TRUQUE DO ABOUT:BLANK) ---
                        # Força o motor a soltar a pasta atual da memória RAM antes de morrer
                        $browserToKill.Navigate("about:blank")
                        [System.Windows.Forms.Application]::DoEvents() # Dá tempo ao Windows para processar
                        
                        $browserToKill.Dispose() 
                    }
                    $tabToClose.Dispose()
                } catch {}

                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                # ====================================================================

                if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
                return
            }
        }
    }.GetNewClosure())
# --- ARRASTAR E SOLTAR ABAS ---
    $tabControl.AllowDrop = $true
    $indicatorColor = [System.Drawing.ColorTranslator]::FromHtml("#FF8C00")

    $dropMarker = New-Object System.Windows.Forms.Panel
    $dropMarker.BackColor = $indicatorColor
    $dropMarker.Width = 3
    $dropMarker.Enabled = $false
    $dropMarker.Visible = $false
    $panel.Controls.Add($dropMarker)

    $arrowMarker = New-Object System.Windows.Forms.Panel
    $arrowMarker.BackColor = $indicatorColor
    $arrowMarker.Size = New-Object System.Drawing.Size(13, 7)
    $arrowMarker.Enabled = $false
    $arrowMarker.Visible = $false

    $arrowPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $arrowPath.AddPolygon([System.Drawing.Point[]]@(
        [System.Drawing.Point]::new(6, 0),
        [System.Drawing.Point]::new(13, 7),
        [System.Drawing.Point]::new(0, 7)
    ))
    $arrowMarker.Region = New-Object System.Drawing.Region($arrowPath)
    
    # ==========================================================
    # --- CORREÇÃO DE VAZAMENTO GDI (Destrói a forma do molde) ---
    # ==========================================================
    $arrowPath.Dispose() 

    $panel.Controls.Add($arrowMarker)

    $ghostForm = New-Object System.Windows.Forms.Form
    $ghostForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $ghostForm.ShowInTaskbar = $false
    $ghostForm.TopMost = $true
    $ghostForm.Enabled = $false
    $ghostForm.Opacity = 0.80
    $ghostForm.BackColor = [System.Drawing.Color]::Magenta
    $ghostForm.TransparencyKey = [System.Drawing.Color]::Magenta
    
    $ghostPB = New-Object System.Windows.Forms.PictureBox
    $ghostPB.BackColor = [System.Drawing.Color]::Magenta
    $ghostPB.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ghostForm.Controls.Add($ghostPB)

    $ghostTimer = New-Object System.Windows.Forms.Timer
    $ghostTimer.Interval = 10 
    $ghostTimer.Add_Tick({
        $bag = $this.Tag
        if ($bag -ne $null) {
            # O X acompanha o local exato do clique para manter a fluidez natural
            $mouseX = [int][System.Windows.Forms.Cursor]::Position.X - $bag.OffsetX
            
            # O Y fica cravado logo ABAIXO da ponta do mouse. A agulha do mouse fica livre!
            $mouseY = [int][System.Windows.Forms.Cursor]::Position.Y + 2 
            
            $ghostForm.Location = New-Object System.Drawing.Point($mouseX, $mouseY)
        }
    }.GetNewClosure())

    $tabControl.Add_MouseMove({ param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $s.Tag.DragStart) {
            $dragStart = $s.Tag.DragStart
            
            $dragSize = [System.Windows.Forms.SystemInformation]::DragSize
            $dragRect = [System.Drawing.Rectangle]::new(
                $dragStart.X - ($dragSize.Width / 2),
                $dragStart.Y - ($dragSize.Height / 2),
                $dragSize.Width,
                $dragSize.Height
            )

            if (-not $dragRect.Contains($e.Location)) {
                $s.Tag.DragStart = $null 

                for ($i = 0; $i -lt $s.TabPages.Count; $i++) {
                    $r = $s.GetTabRect($i)
                    
                    $boxSize = 14
                    $boxX = $r.Right - 18
                    $boxY = $r.Top + [int](($r.Height - $boxSize) / 2)
                    $closeRect = [System.Drawing.Rectangle]::new($boxX, $boxY, $boxSize, $boxSize)
                    
                    if ($r.Contains($dragStart) -and -not $closeRect.Contains($dragStart)) { 
                        
                        $offsetX = $dragStart.X - $r.X
                        $ghostTimer.Tag = @{ OffsetX = $offsetX }

                        $bmp = New-Object System.Drawing.Bitmap($r.Width, $r.Height)
                        $g = [System.Drawing.Graphics]::FromImage($bmp)
                        $g.Clear([System.Drawing.Color]::Magenta) 
                        
                        $bgBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(80, 80, 85))
                        $g.FillRectangle($bgBrush, 0, 0, $r.Width, $r.Height)
                        $g.DrawString($s.TabPages[$i].Text, $s.Font, [System.Drawing.Brushes]::White, 5, 5)
                        
                        $bgBrush.Dispose()
                        $g.Dispose()
                        
                        $ghostPB.Image = $bmp
                        $ghostForm.Size = $bmp.Size
                        
                        $ghostForm.Location = New-Object System.Drawing.Point(
                            ([System.Windows.Forms.Cursor]::Position.X - $offsetX),
                            ([System.Windows.Forms.Cursor]::Position.Y + 2)
                        )
                        
                        $ghostForm.Show()
                        $ghostTimer.Start()

                        # ==============================================================
                        # A ANESTESIA CORRIGIDA: Desliga a "caixa" (Parent) do navegador
                        # ==============================================================
                        $leftB = if ($global:LeftTabControl -and $global:LeftTabControl.SelectedTab) { $global:LeftTabControl.SelectedTab.Controls[0] } else { $null }
                        $rightB = if ($global:RightTabControl -and $global:RightTabControl.SelectedTab) { $global:RightTabControl.SelectedTab.Controls[0] } else { $null }
                        
                        # Desliga a Aba (Parent) para bloquear os cliques e reações
                        if ($leftB -and $leftB.Parent) { $leftB.Parent.Enabled = $false }
                        if ($rightB -and $rightB.Parent) { $rightB.Parent.Enabled = $false }

                        # O script pausa aqui e faz o arrasto
                        $s.DoDragDrop($s.TabPages[$i], [System.Windows.Forms.DragDropEffects]::Move) | Out-Null
                        
                        # O arrasto terminou, acorda as abas instantaneamente
                        if ($leftB -and $leftB.Parent) { $leftB.Parent.Enabled = $true }
                        if ($rightB -and $rightB.Parent) { $rightB.Parent.Enabled = $true }
                        # ==============================================================
                        
                        $ghostTimer.Stop()
                        $ghostForm.Hide()
                        
                        # --- CORREÇÃO DE CRASH: Libera a imagem do PictureBox antes de destruir a original ---
                        $ghostPB.Image = $null
                        $bmp.Dispose()
                        
                        return
                    }
                }
            }
        }
    }.GetNewClosure())

    $tabControl.Add_GiveFeedback({ param($s, $e)
        $e.UseDefaultCursors = $false
        if ($e.Effect -eq [System.Windows.Forms.DragDropEffects]::Move) {
            # Sobre a barra de abas (área permitida)
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        } else {
            # Fora da barra (sobre os arquivos, outro painel, etc)
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::No
        }
    }.GetNewClosure())

    $tabControl.Add_DragOver({ param($s, $e)
        # 1. Se estiver arrastando uma ABA (Lógica Antiga)
        if ($e.Data.GetDataPresent([System.Windows.Forms.TabPage])) {
            $clientPoint = $s.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
            
            # --- FRONTEIRA ABSOLUTA: A linha exata que divide abas e conteúdo ---
            $boundaryY = $s.DisplayRectangle.Top
            
            # Se o mouse descer para a área dos arquivos, bloqueia visualmente
            if ($clientPoint.Y -ge $boundaryY) {
                $e.Effect = [System.Windows.Forms.DragDropEffects]::None
                $dropMarker.Visible = $false
                $arrowMarker.Visible = $false
                return
            }

            $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
            $found = $false
            
            for ($i = 0; $i -lt $s.TabPages.Count; $i++) {
                $rect = $s.GetTabRect($i)
                # MÁGICA: Verifica APENAS a mira horizontal (X), deixando o arraste muito mais fácil
                if ($clientPoint.X -ge $rect.Left -and $clientPoint.X -le $rect.Right) {
                    $isRightHalf = ($clientPoint.X - $rect.Left) -gt ($rect.Width / 2)
                    
                    $dropMarker.Height = $rect.Height
                    $markerY = $s.Top + $rect.Top
                    
                    if ($isRightHalf) { $markerX = $s.Left + $rect.Right } else { $markerX = $s.Left + $rect.Left }
                    
                    $dropMarker.Location = New-Object System.Drawing.Point($markerX, $markerY)
                    
                    $arrowX = $markerX - 5
                    $arrowY = $markerY + $rect.Height
                    $arrowMarker.Location = New-Object System.Drawing.Point($arrowX, $arrowY)
                    
                    $dropMarker.Visible = $true
                    $arrowMarker.Visible = $true
                    $dropMarker.BringToFront()
                    $arrowMarker.BringToFront()
                    $found = $true
                    break
                }
            }
            
            # Se não encontrou nenhuma aba embaixo do mouse, joga para o espaço vazio à direita
            if (-not $found -and $s.TabPages.Count -gt 0) {
                $lastRect = $s.GetTabRect($s.TabPages.Count - 1)
                # Qualquer espaço além da última aba é válido
                if ($clientPoint.X -gt $lastRect.Right) {
                    $markerX = $s.Left + $lastRect.Right
                    $markerY = $s.Top + $lastRect.Top
                    
                    $dropMarker.Height = $lastRect.Height
                    $dropMarker.Location = New-Object System.Drawing.Point($markerX, $markerY)
                    
                    $arrowX = $markerX - 5
                    $arrowY = $markerY + $lastRect.Height
                    $arrowMarker.Location = New-Object System.Drawing.Point($arrowX, $arrowY)
                    
                    $dropMarker.Visible = $true
                    $arrowMarker.Visible = $true
                    $dropMarker.BringToFront()
                    $arrowMarker.BringToFront()
                    $found = $true
                }
            }
            
            if (-not $found) { 
                $e.Effect = [System.Windows.Forms.DragDropEffects]::None
                $dropMarker.Visible = $false
                $arrowMarker.Visible = $false 
            }
            
        # 2. NOVO: Se estiver arrastando uma PASTA de fora (Windows Explorer)
        } elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy # Cursor de cópia (+)
            $dropMarker.Visible = $false
            $arrowMarker.Visible = $false
        } else {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::None
            $dropMarker.Visible = $false
            $arrowMarker.Visible = $false
        }
    }.GetNewClosure())

    $tabControl.Add_DragLeave({ param($s, $e)
        $dropMarker.Visible = $false
        $arrowMarker.Visible = $false
    }.GetNewClosure())

    $tabControl.Add_DragDrop({ param($s, $e)
        $dropMarker.Visible = $false 
        $arrowMarker.Visible = $false
        
        # 1. Se soltou uma ABA (Reordenação)
        if ($e.Data.GetDataPresent([System.Windows.Forms.TabPage])) {
            $draggedTab = $e.Data.GetData([System.Windows.Forms.TabPage])
            if ($draggedTab -ne $null) {
                $clientPoint = $s.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
                $targetIndex = -1
                
                for ($i = 0; $i -lt $s.TabPages.Count; $i++) {
                    $rect = $s.GetTabRect($i)
                    # Verifica apenas o X, acompanhando a melhoria do DragOver
                    if ($clientPoint.X -ge $rect.Left -and $clientPoint.X -le $rect.Right) {
                        $isRightHalf = ($clientPoint.X - $rect.Left) -gt ($rect.Width / 2)
                        if ($isRightHalf) { $targetIndex = $i + 1 } else { $targetIndex = $i }
                        break
                    }
                }
                
                if ($targetIndex -eq -1 -and $s.TabPages.Count -gt 0) {
                    $lastRect = $s.GetTabRect($s.TabPages.Count - 1)
                    # Qualquer clique no lado direito manda a aba para o final
                    if ($clientPoint.X -gt $lastRect.Right) {
                        $targetIndex = $s.TabPages.Count
                    }
                }
                
                if ($targetIndex -ne -1) {
                    $currentIndex = $s.TabPages.IndexOf($draggedTab)
                    if ($targetIndex -gt $currentIndex) { $targetIndex-- }
                    
                    if ($targetIndex -ne $currentIndex) {
                        $s.TabPages.Remove($draggedTab)
                        $s.TabPages.Insert($targetIndex, $draggedTab)
                        $s.SelectedTab = $draggedTab
                    }
                }
                
                # --- GATILHO RECOVERY: Salva a nova ordem após soltar a aba ---
                if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
            }
            
        # 2. NOVO: Se soltou uma PASTA externa
        } elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $droppedFiles = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
            if ($droppedFiles -and $droppedFiles.Count -gt 0) {
                $targetPath = $droppedFiles[0] # Pega o primeiro item que você soltou
                
                # CORREÇÃO: Usando -LiteralPath para aceitar pastas com [colchetes] e nomes complexos
                if (Test-Path -LiteralPath $targetPath -PathType Container) {
                    # Ativa a função mágica invisível de criar nova aba!
                    & ($s.Tag.AddTab) -TargetTabControl $s -Path $targetPath -AddressBox ($s.Tag.TxtPath) -bBack ($s.Tag.BtnBack) -bFwd ($s.Tag.BtnFwd) -BtnView ($s.Tag.BtnView)
                    if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
                }
            }
        }
    }.GetNewClosure())

    # --- MENU DE CONTEXTO DAS ABAS (CLIQUE DIREITO) ---
    $tabContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $itemDup = $tabContextMenu.Items.Add("Duplicar Aba")
    $itemDup.Add_Click({
        $tInfo = $tabContextMenu.Tag
        if ($tInfo.Path) { & ($tInfo.TC.Tag.AddTab) -TargetTabControl $tInfo.TC -Path $tInfo.Path -AddressBox ($tInfo.TC.Tag.TxtPath) -bBack ($tInfo.TC.Tag.BtnBack) -bFwd ($tInfo.TC.Tag.BtnFwd) -BtnView ($tInfo.TC.Tag.BtnView) }
        
        # --- GATILHO RECOVERY: Salva ao Duplicar ---
        if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
    }.GetNewClosure())

    $itemCopy = $tabContextMenu.Items.Add("Copiar para o outro painel")
    $itemCopy.Add_Click({
        $tInfo = $tabContextMenu.Tag
        $otherTC = if ($tInfo.TC -eq $global:LeftTabControl) { $global:RightTabControl } else { $global:LeftTabControl }
        if ($otherTC -and $tInfo.Path) {
            $otherUI = $otherTC.Tag
            & ($otherUI.AddTab) -TargetTabControl $otherTC -Path $tInfo.Path -AddressBox $otherUI.TxtPath -bBack $otherUI.BtnBack -bFwd $otherUI.BtnFwd -BtnView $otherUI.BtnView
            
            # --- GATILHO RECOVERY: Salva ao Copiar ---
            if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
        }
    }.GetNewClosure())

    $itemMove = $tabContextMenu.Items.Add("Mover para o outro painel")
    $itemMove.Add_Click({
        $tInfo = $tabContextMenu.Tag
        $otherTC = if ($tInfo.TC -eq $global:LeftTabControl) { $global:RightTabControl } else { $global:LeftTabControl }
        if ($otherTC -and $tInfo.Path) {
            $otherUI = $otherTC.Tag
            & ($otherUI.AddTab) -TargetTabControl $otherTC -Path $tInfo.Path -AddressBox $otherUI.TxtPath -bBack $otherUI.BtnBack -bFwd $otherUI.BtnFwd -BtnView $otherUI.BtnView
            
            $tInfo.TC.TabPages.Remove($tInfo.Tab)
            if ($tInfo.TC.TabPages.Count -eq 0) { 
                & ($tInfo.TC.Tag.AddTab) -TargetTabControl $tInfo.TC -Path "C:\" -AddressBox ($tInfo.TC.Tag.TxtPath) -bBack ($tInfo.TC.Tag.BtnBack) -bFwd ($tInfo.TC.Tag.BtnFwd) -BtnView ($tInfo.TC.Tag.BtnView) 
            }
            
            # --- GATILHO RECOVERY: Salva ao Mover ---
            if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
        }
    }.GetNewClosure())
    
    [void]$tabContextMenu.Items.Add("-")

    $itemClose = $tabContextMenu.Items.Add("Fechar Aba")
    $itemClose.Add_Click({
        $tInfo = $tabContextMenu.Tag
        
        # ====================================================================
        # LIXEIRO SEGURO (Menu de Contexto)
        # ====================================================================
        $tabToClose = $tInfo.Tab
        $currentIndex = $tInfo.TC.TabPages.IndexOf($tabToClose)

        # MÁGICA DE NAVEGADOR: Pula para a aba vizinha ANTES de deletar a atual
        if ($tInfo.TC.TabPages.Count -gt 1 -and $tInfo.TC.SelectedTab -eq $tabToClose) {
            if ($currentIndex -eq ($tInfo.TC.TabPages.Count - 1)) {
                $tInfo.TC.SelectedIndex = $currentIndex - 1 # Se era a última, foca na da esquerda
            } else {
                $tInfo.TC.SelectedIndex = $currentIndex + 1 # Se não, foca na da direita
            }
        }
        
        # 1. Tira a aba da tela PRIMEIRO para a aba vizinha assumir com segurança
        $tInfo.TC.TabPages.Remove($tabToClose) 
        
        # 2. Se era a última aba, recria o "Meu Computador" para não ficar vazio
        if ($tInfo.TC.TabPages.Count -eq 0) { 
            & ($tInfo.TC.Tag.AddTab) -TargetTabControl $tInfo.TC -Path "C:\" -AddressBox ($tInfo.TC.Tag.TxtPath) -bBack ($tInfo.TC.Tag.BtnBack) -bFwd ($tInfo.TC.Tag.BtnFwd) -BtnView ($tInfo.TC.Tag.BtnView) 
        }

        # 3. Destrói o navegador nos bastidores de forma limpa
        try {
            if ($tabToClose.Controls.Count -gt 0) {
                $browserToKill = $tabToClose.Controls[0]
                if ($global:ActiveBrowser -eq $browserToKill) { $global:ActiveBrowser = $null }
                
                # --- CORRECAO DE MEMORY LEAK (O TRUQUE DO ABOUT:BLANK) ---
                $browserToKill.Navigate("about:blank")
                [System.Windows.Forms.Application]::DoEvents() 
                
                $browserToKill.Dispose()
            }
            $tabToClose.Dispose()
        } catch {}

        # 4. Passa o caminhão de lixo da memória RAM
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        # ====================================================================

        # --- GATILHO RECOVERY: Salva ao Fechar ---
        if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
    }.GetNewClosure())

    $tabControl.Add_MouseUp({ param($s, $e)
        $s.Tag.DragStart = $null # Limpa o gatilho de arrasto ao soltar o clique
        
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            for ($i = 0; $i -lt $s.TabPages.Count; $i++) {
                $r = $s.GetTabRect($i)
                if ($r.Contains($e.Location)) {
                    $clickedTab = $s.TabPages[$i]
                    $browser = $clickedTab.Controls[0]
                    $path = "C:\"
                    try { if ($browser.Url) { $path = $browser.Url.LocalPath } } catch {}
                    
                    $tabContextMenu.Tag = @{ TC = $s; Tab = $clickedTab; Path = $path }
                    $tabContextMenu.Show($s, $e.Location)
                    return
                }
            }
        }
    }.GetNewClosure())

    $tabControl.Add_SelectedIndexChanged({
        if ($tabControl.SelectedTab -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $currBrowser = $tabControl.SelectedTab.Controls[0]
            if ($currBrowser) {
                $global:ActiveBrowser = $currBrowser
                if ($currBrowser.Url) { $txtPath.Text = $currBrowser.Url.LocalPath }
                $btnBack.Enabled = $currBrowser.CanGoBack; $btnFwd.Enabled = $currBrowser.CanGoForward
            }
        }
        $tabControl.Invalidate()
    }.GetNewClosure())

    # --- EVENTOS DOS BOTÕES ---
    $btnBack.Add_Click({ 
        if ($tabControl.SelectedTab -ne $null -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]; if($b.CanGoBack){$b.GoBack()} 
        }
    }.GetNewClosure())

    $btnFwd.Add_Click({ 
        if ($tabControl.SelectedTab -ne $null -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]; if($b.CanGoForward){$b.GoForward()} 
        }
    }.GetNewClosure())

    $btnUp.Add_Click({ 
        if ($tabControl.SelectedTab -ne $null -and $tabControl.SelectedTab.Controls.Count -gt 0) {
            $b = $tabControl.SelectedTab.Controls[0]
            try { 
                if ($b.Url) { 
                    $currentPath = $b.Url.LocalPath
                    $parentPath = Split-Path $currentPath -Parent
                    
                    if ($parentPath) { 
                        # A INJEÇÃO DE INTELIGÊNCIA: Guarda o nome da pasta atual antes de subir!
                        $folderToSelect = Split-Path $currentPath -Leaf
                        if ($null -eq $global:FolderSelMemory) { $global:FolderSelMemory = @{} }
                        
                        # Subscreve a memória da pasta pai com o nome da pasta que estamos abandonando
                        $global:FolderSelMemory[$parentPath] = @($folderToSelect)
                        
                        $b.Navigate($parentPath) 
                    } 
                } 
            } catch {} 
        }
    }.GetNewClosure())

    $txtPath.Add_KeyDown({ param($s, $e) 
        # 1. NAVEGAR (Pressionar Enter)
        if($e.KeyCode -eq 'Enter'){ 
            if ($tabControl.SelectedTab -ne $null -and $tabControl.SelectedTab.Controls.Count -gt 0) {
                $b = $tabControl.SelectedTab.Controls[0]; $b.Navigate($txtPath.Text)
                $e.SuppressKeyPress=$true 
            }
        } 
        # 2. SELECIONAR TUDO (Ctrl + A)
        elseif ($e.Control -and $e.KeyCode -eq 'A') {
            $s.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
        # 3. APAGAR PALAVRA ANTERIOR (Ctrl + Backspace)
        elseif ($e.Control -and $e.KeyCode -eq 'Back') {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            if ($s.SelectionStart -gt 0) {
                # Procura a última barra (\ ou /) ou espaço para saber onde a palavra/pasta termina
                $stopIndex = $s.Text.LastIndexOfAny(@('\', '/', ' '), $s.SelectionStart - 2)
                if ($stopIndex -lt 0) { $stopIndex = 0 } else { $stopIndex++ }
                
                $s.Text = $s.Text.Remove($stopIndex, $s.SelectionStart - $stopIndex)
                $s.SelectionStart = $stopIndex
            }
        }
        # 4. DESFAZER (Ctrl + Z) E REFAZER (Ctrl + Y)
        elseif ($e.Control -and ($e.KeyCode -eq 'Z' -or $e.KeyCode -eq 'Y')) {
            # O WinForms usa o Undo como um interruptor (Desfaz/Refaz a última ação)
            if ($s.CanUndo) { $s.Undo() }
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    }.GetNewClosure())

    # --- CLIQUE NA ESTRELA (ADICIONAR/REMOVER FAVORITO) ---
    $btnFav.Add_Click({
        if ($tabControl.SelectedTab -eq $null -or $tabControl.SelectedTab.Controls.Count -eq 0) { return }
        $currBrowser = $tabControl.SelectedTab.Controls[0]
        try {
            if (-not $currBrowser.Url) { return }
            $currentPath = $currBrowser.Url.LocalPath
            $currentName = Split-Path $currentPath -Leaf; if ([string]::IsNullOrEmpty($currentName)) { $currentName = $currentPath }
            $data = Get-BookmarksData
            $isFav = Test-PathIsFavorite -TargetData $data -SearchPath $currentPath
            
            if ($isFav) {
                if ([System.Windows.Forms.MessageBox]::Show("Remover '$currentName' dos favoritos?", "Remover Favorito", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes') {
                    Remove-FavoriteByPath -PathToRemove $currentPath
                    if ($global:StarEmptyIconBmp) { $btnFav.Image = $global:StarEmptyIconBmp }
                }
            } else {
                $choice = [System.Windows.Forms.MessageBox]::Show("Deseja escolher a pasta de destino?", "Novo Favorito", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($choice -eq 'Yes') {
                    $result = Show-FolderSelectionDialog
                    if ($result) { 
                        Add-Favorite -Name $currentName -Path $currentPath -ParentNodeData $result.Selected
                        Save-BookmarksData $result.Root
                        if ($global:StarIconBmp) { $btnFav.Image = $global:StarIconBmp } 
                    }
                } else { 
                    Add-Favorite -Name $currentName -Path $currentPath -ParentNodeData $null
                    if ($global:StarIconBmp) { $btnFav.Image = $global:StarIconBmp } 
                }
            }
        } catch {}
    }.GetNewClosure())

    # --- MONITOR TIMER ---
    $monitorTimer = New-Object System.Windows.Forms.Timer; $monitorTimer.Interval = 300 
    
    $monitorTimer.Tag = @{ TabCtrl = $tabControl; StatusPanel = $statusPanel; Label = $lblStatus; DiskLabel = $lblDisk; FavBtn = $btnFav; LastSig = ""; BtnAdd = $btnNewTab; IndLeft = $lblMoreLeft; TxtPath = $txtPath; LastDiskSig = ""; BtnDel = $btnDelete; BtnBkp = $btnBackup; BtnCut = $btnCut; BtnCopy = $btnCopyFile; BtnPaste = $btnPaste; BtnProp = $btnProp; PrevForm = $previewForm; PrevPB = $pbPreview; PrevRTB = $rtbPreview; PrevLbl = $lblPreviewMsg; PrevWMP = $wmpPreview; PrevWV = $wv2Preview; PrevGif = $wfGifPreview; PrevStatus = $lblPreviewStatus; LastPrevPath = ""; LastPrevWriteTime = [DateTime]::MinValue; LastSelPath = ""; LastSelNeighbor = ""; AppForm = $form; BootTicks = 0 }
    $monitorTimer.Tag["SwitchProf"] = $SwitchProfile
    $monitorTimer.Add_Tick({
        $bag = $this.Tag
        $bag.BootTicks++ 
        
        try {
            $hideBtnAdd = $false
            $hideIndLeft = $true 

            # ========================================================================
            # A BARREIRA INVISIVEL (PAREDE ESQUERDA) - Mantida
            # ========================================================================
            if ($null -eq $bag.TabCtrl.Tag.WallLeft) {
                $wall = New-Object System.Windows.Forms.Panel
                $wall.Size = New-Object System.Drawing.Size(2, 22) 
                if ($bag.TabCtrl.Parent) { $wall.BackColor = $bag.TabCtrl.Parent.BackColor }
                else { $wall.BackColor = [System.Drawing.SystemColors]::Control }
                $bag.TabCtrl.Parent.Controls.Add($wall)
                $bag.TabCtrl.Tag.WallLeft = $wall
            }
            # ========================================================================

            if ($bag.TabCtrl.TabCount -gt 0 -and $bag.BootTicks -gt 3) {
                try {
                    $firstRect = $bag.TabCtrl.GetTabRect(0)
                    if ($firstRect.X -lt 0) { $hideIndLeft = $false }
                } catch {}

                # ========================================================================
                # LADO DIREITO REVERTIDO PARA O SEU ORIGINAL (40 pixels)
                # ========================================================================
                for ($i = 0; $i -lt $bag.TabCtrl.TabCount; $i++) {
                    try {
                        $r = $bag.TabCtrl.GetTabRect($i)
                        if ($r.Width -eq 0 -or $r.X -lt 0 -or $r.Right -gt ($bag.TabCtrl.Width - 40)) {
                            $hideBtnAdd = $true
                            break
                        }
                    } catch {}
                }
            }
            
            # ========================================================================
            # CONTROLE DO INDICADOR [<<] E BARREIRA ESQUERDA - Mantido
            # ========================================================================
            if ($hideIndLeft) {
                if ($bag.IndLeft.Visible) { $bag.IndLeft.Visible = $false }
                if ($bag.TabCtrl.Tag.WallLeft -ne $null) { $bag.TabCtrl.Tag.WallLeft.Visible = $false }
            } else {
                if (-not $bag.IndLeft.Visible) { $bag.IndLeft.Visible = $true }
                
                $tabHeight = 22
                try { $tabHeight = $bag.TabCtrl.GetTabRect(0).Height + 3 } catch {}
                
                $bag.IndLeft.AutoSize = $false
                $bag.IndLeft.Height = $tabHeight
                $bag.IndLeft.Top = $bag.TabCtrl.Top
                $bag.IndLeft.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                
                $bag.TabCtrl.Tag.WallLeft.Size = New-Object System.Drawing.Size(2, $tabHeight)
                $bag.TabCtrl.Tag.WallLeft.Location = New-Object System.Drawing.Point($bag.IndLeft.Right, $bag.TabCtrl.Top)
                $bag.TabCtrl.Tag.WallLeft.Visible = $true
                $bag.TabCtrl.Tag.WallLeft.BringToFront()
                $bag.IndLeft.BringToFront()
            }

            # Aplicacao da visibilidade do seu botao original
            if ($hideBtnAdd) {
                if ($bag.BtnAdd.Visible) { $bag.BtnAdd.Visible = $false }
            } else {
                if (-not $bag.BtnAdd.Visible) { $bag.BtnAdd.Visible = $true }
            }

            if ($bag.TabCtrl.SelectedTab -ne $null -and $bag.TabCtrl.SelectedTab.Controls.Count -gt 0) {
                $currBrowser = $bag.TabCtrl.SelectedTab.Controls[0]
                if ($currBrowser.Focused) { $global:ActiveBrowser = $currBrowser }
                
                $shellView = $currBrowser.ActiveXInstance.Document
                
                # --- MODIFICAÇÃO SEGURA: Desmembrando variáveis para Lixeira ---
                if ($shellView) {
                    $folder = $shellView.Folder
                    if ($folder) {
                        $selfObj = $folder.Self
                        if ($selfObj) {
                            $rawPath = $selfObj.Path
                            
                            # --- OPÇÃO 1A: Sobrevivência se a pasta for deletada ou Pendrive ejetado ---
                            $isVirt = ($rawPath -match "^(shell|::|search\-ms)")
                            if (-not $isVirt -and -not [string]::IsNullOrWhiteSpace($rawPath)) {
                                
                                # Monitora ativamente o Hardware
                                $isDriveReady = $true
                                if ($rawPath -match "^([A-Za-z]:)") {
                                    try { $isDriveReady = (New-Object System.IO.DriveInfo($matches[1])).IsReady } catch { $isDriveReady = $false }
                                }

                                if (-not $isDriveReady -or -not (Test-Path -LiteralPath $rawPath -ErrorAction SilentlyContinue)) {
                                    # A pasta atual sumiu ou o disco fisico foi puxado!
                                    try {
                                        if (-not $isDriveReady) {
                                            $currBrowser.Navigate("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
                                        } else {
                                            $parentPath = Split-Path $rawPath -Parent -ErrorAction Stop
                                            if (-not [string]::IsNullOrWhiteSpace($parentPath) -and (Test-Path -LiteralPath $parentPath)) { $currBrowser.Navigate($parentPath) }
                                            else { $currBrowser.Navigate("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") }
                                        }
                                    } catch { $currBrowser.Navigate("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") }
                                    
                                    # --- MICRO-LEAK CORRIGIDO AQUI ANTES DO RETURN ---
                                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selfObj) | Out-Null
                                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null
                                    return # Interrompe a leitura deste milissegundo
                                }
                            }

                            $friendlyName = $folder.Title
                            
                            $displayPath = $rawPath
                            if ($rawPath -match "^::" -or [string]::IsNullOrWhiteSpace($rawPath)) { $displayPath = $friendlyName }
                            if ([string]::IsNullOrWhiteSpace($displayPath)) { $displayPath = $currBrowser.ActiveXInstance.LocationName }

                            if (-not $bag.TxtPath.Focused -and $bag.TxtPath.Text -ne $displayPath -and -not [string]::IsNullOrWhiteSpace($displayPath)) {
                                $bag.TxtPath.Text = $displayPath
                            }

                            $tabName = $displayPath
                            if ($displayPath -match "\\") { try { $tabName = Split-Path $displayPath -Leaf } catch {} }
                            if ([string]::IsNullOrWhiteSpace($tabName)) { $tabName = $displayPath }
                            if ($tabName.Length -gt 21) { $tabName = $tabName.Substring(0, 21) + "..." }
                            
                            $expectedTabText = "$tabName      "
                            if ($bag.TabCtrl.SelectedTab.Text -ne $expectedTabText) {
                                $bag.TabCtrl.SelectedTab.Text = $expectedTabText
                            }
                        
                            $curPath = $bag.TxtPath.Text
                            if (-not [string]::IsNullOrWhiteSpace($curPath)) {
                                $data = Get-BookmarksData
                                if (Test-PathIsFavorite -TargetData $data -SearchPath $curPath) {
                                    if ($global:StarIconBmp -and $bag.FavBtn.Image -ne $global:StarIconBmp) { $bag.FavBtn.Image = $global:StarIconBmp }
                                } else {
                                    if ($global:StarEmptyIconBmp -and $bag.FavBtn.Image -ne $global:StarEmptyIconBmp) { $bag.FavBtn.Image = $global:StarEmptyIconBmp }
                                }
                            }
                            
                            $totalItems = 0; try { $totalItems = $folder.Items().Count } catch {} 
                            $items = $shellView.SelectedItems(); $selCount = $items.Count
                            
                            # --- INICIO DA MEMORIA DE FOCO (OTIMIZADA) ---
                            if ($selCount -eq 1) {
                                $currentItem = $items.Item(0)
                                $currentPath = $currentItem.Path
                                
                                # O SEGREDO DA VELOCIDADE: So executa a leitura pesada SE a selecao for nova!
                                if ($currentPath -ne $bag.LastSelPath) {
                                    $bag.LastSelPath = $currentPath
                                    $bag.LastSelNeighbor = ""
                                    
                                    # Espiao: Guarda o nome do vizinho
                                    $allItems = $folder.Items()
                                    for ($i = 0; $i -lt $allItems.Count; $i++) {
                                        if ($allItems.Item($i).Path -eq $currentPath) {
                                            if ($i -gt 0) {
                                                $bag.LastSelNeighbor = $allItems.Item($i - 1).Name
                                            }
                                            break
                                        }
                                    }
                                }
                                
                                # --- CORREÇÃO DE MICRO-LEAK 1 ---
                                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($currentItem) | Out-Null
                                
                            } elseif ($selCount -eq 0 -and $bag.LastSelPath -ne "") {
                                # Nada selecionado! Vamos investigar se o arquivo foi apagado do disco
                                if (-not (Test-Path -LiteralPath $bag.LastSelPath)) {
                                    # Confirmado: O arquivo sumiu. Vamos resgatar o vizinho de cima!
                                    if ($bag.LastSelNeighbor -ne "") {
                                        $neighborItem = $folder.ParseName($bag.LastSelNeighbor)
                                        if ($neighborItem) {
                                            # Comando 29 = Selecionar e focar
                                            try { $shellView.SelectItem($neighborItem, 29) } catch {}
                                            
                                            # --- CORREÇÃO DE MICRO-LEAK 2 ---
                                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($neighborItem) | Out-Null
                                        }
                                    }
                                }
                                # Limpa a memoria para encerrar o ciclo
                                $bag.LastSelPath = ""
                                $bag.LastSelNeighbor = ""
                            }
                            # --- FIM DA MEMORIA DE FOCO ---

                            if ($null -eq $global:FolderSelMemory) { $global:FolderSelMemory = @{}; $global:BrowserPath = @{} }
                            $bId = $currBrowser.Handle.ToString()

                            if ($global:BrowserPath[$bId] -ne $rawPath) {
                                $global:BrowserPath[$bId] = $rawPath
                                if ($global:FolderSelMemory.ContainsKey($rawPath)) {
                                    $savedNames = $global:FolderSelMemory[$rawPath]
                                    if ($savedNames.Count -gt 0) {
                                        foreach ($name in $savedNames) {
                                            try {
                                                $tgt = $folder.ParseName($name)
                                                if ($tgt) { 
                                                    $shellView.SelectItem($tgt, 17) 
                                                    
                                                    # --- CORREÇÃO DE MICRO-LEAK 3 ---
                                                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tgt) | Out-Null
                                                } 
                                            } catch {}
                                        }
                                        
                                        try { 
                                            $currBrowser.Select()
                                            $parentForm = $currBrowser.FindForm()
                                            if ($parentForm -ne $null) { 
                                                $parentForm.ActiveControl = $currBrowser 
                                            }
                                        } catch {}

                                        $items = $shellView.SelectedItems(); $selCount = $items.Count
                                    }
                                }
                            } else {
                                # A MÁGICA: Só atualiza a memória se o mouse estiver ativamente na pasta!
                                # Se o usuário clicar num botão/menu, a memória congela e protege a seleção.
                                if ($currBrowser.Focused) {
                                    if ($totalItems -gt 0) {
                                        $memArr = @()
                                        if ($selCount -gt 0) { foreach ($i in $items) { try { $memArr += $i.Name } catch {} } }
                                        $global:FolderSelMemory[$rawPath] = $memArr
                                    }
                                }
                            }

                            $urlStr = ""
                            try { if ($currBrowser.Url) { $urlStr = $currBrowser.Url.ToString() } } catch {}
                            $sig = "T:$totalItems|S:$selCount|$urlStr"; $totalSize = 0
                            
                            try {
                                $isVirtual = ($rawPath -match "^::")
                                $isRecycleBin = ($rawPath -match "645FF040-5081-101B-9F08-00AA002F954E")
                                $isMyComputer = ($rawPath -match "20D04FE0-3AEA-1069-A2D8-08002B30309D")

                                if ($bag.BtnCut -ne $null) { $bag.BtnCut.Enabled = [bool]($selCount -gt 0 -and -not $isVirtual) }
                                if ($bag.BtnCopy -ne $null) { $bag.BtnCopy.Enabled = [bool]($selCount -gt 0 -and (-not $isVirtual -or $isMyComputer)) }
                                if ($bag.BtnPaste -ne $null) { $bag.BtnPaste.Enabled = [bool](-not $isVirtual) }
                                if ($bag.BtnDel -ne $null) { $bag.BtnDel.Enabled = [bool]($selCount -gt 0 -and (-not $isVirtual -or $isRecycleBin)) }
                                if ($bag.BtnBkp -ne $null) { $bag.BtnBkp.Enabled = [bool]($selCount -gt 0 -and -not $isVirtual) }
                                if ($bag.BtnProp -ne $null) { $bag.BtnProp.Enabled = [bool]($selCount -gt 0) }
                            } catch {}

                            if ($selCount -gt 0) { 
                                foreach($i in $items) { 
                                    if($i.Path){ $sig += $i.Path.Length; try { $totalSize += $i.Size } catch {} } 
                                } 
                            }
                            
                            if ($sig -ne $bag.LastSig) {
                                $bag.LastSig = $sig
                                $statusText = "  $totalItems itens"
                                if ($selCount -gt 0) {
                                    $sizeStr = Format-FileSize -Bytes $totalSize
                                    $statusText += "  |  $selCount item(s) selecionado(s)  |  $sizeStr"
                                    $bag.StatusPanel.BackColor = "#EBF4FF" 
                                } else { 
                                    $bag.StatusPanel.BackColor = "#F5F5F5" 
                                }
                                $bag.Label.Text = $statusText
                            }

                            if ($bag.PrevForm.Visible) {
                                $selItemPath = ""
                                if ($selCount -eq 1) {
                                    try { $selItemPath = $items.Item(0).Path } catch {}
                                }

                                # --- OPÇÃO 2B: Verificação Híbrida de Atualização Externa ---
                                $forcePreviewUpdate = $false
                                $currentWriteTime = [DateTime]::MinValue
                                
                                if (-not [string]::IsNullOrWhiteSpace($selItemPath) -and (Test-Path $selItemPath -PathType Leaf)) {
                                    try { $currentWriteTime = (Get-Item $selItemPath).LastWriteTime } catch {}
                                    
                                    # Só recarrega se o Clone Commander for a janela ativa e a data do arquivo mudou
                                    if ($bag.AppForm.ContainsFocus -and $selItemPath -eq $bag.LastPrevPath -and $currentWriteTime -ne $bag.LastPrevWriteTime) {
                                        $forcePreviewUpdate = $true
                                    }
                                }

                                if ($selItemPath -ne $bag.LastPrevPath -or $forcePreviewUpdate) {
                                    $bag.LastPrevPath = $selItemPath
                                    $bag.LastPrevWriteTime = $currentWriteTime

                                    if ($bag.PrevPB.Child -ne $null -and $bag.PrevPB.Child.Source -ne $null) {
                                        $bag.PrevPB.Child.Source = $null
                                    }
                                    if ($bag.PrevWMP.Visible) {
                                        try {
                                            $ocx = $bag.PrevWMP.GetMediaPlayer()
                                            $ocx.controls.stop()
                                            $ocx.URL = ""
                                        } catch {}
                                    }
                                    if ($bag.PrevWV -ne $null -and $bag.PrevWV.Visible) {
                                        try { $bag.PrevWV.Source = New-Object System.Uri("about:blank") } catch {}
                                    }

                                    $bag.PrevPB.Visible = $false
                                    $bag.PrevRTB.Visible = $false
                                    $bag.PrevWMP.Visible = $false
                                    if ($bag.PrevWV -ne $null) { $bag.PrevWV.Visible = $false }
                                    
                                    # ====================================================================
                                    # 3. NOVO: LIMPEZA ANTI-LEAK DO GIF (Desbloqueia o arquivo no disco)
                                    # ====================================================================
                                    if ($bag.PrevGif.Image -ne $null) {
                                        $bag.PrevGif.Image.Dispose()
                                        $bag.PrevGif.Image = $null
                                    }
                                    if ($bag.PrevGif.Tag -ne $null) {
                                        $bag.PrevGif.Tag.Dispose() # Mata a stream de memória
                                        $bag.PrevGif.Tag = $null
                                    }
                                    $bag.PrevGif.Visible = $false
                                    # ====================================================================

                                    $bag.PrevLbl.Visible = $true
                                    
                                    # --- NOVO: LOGICA DE METADADOS (TAMANHO E RESOLUCAO EM 2º PLANO) ---
                                    # A barra nunca mais fica invisível. Apenas limpamos o texto enquanto o novo carrega.
                                    $bag.PrevStatus.Text = ""
                                    
                                    # 1. Adicionado -LiteralPath no Test-Path
                                    if (-not [string]::IsNullOrWhiteSpace($selItemPath) -and (Test-Path -LiteralPath $selItemPath -PathType Leaf)) {
                                        try {
                                            # 2. Adicionado -LiteralPath no Get-Item
                                            $fileInfo = Get-Item -LiteralPath $selItemPath
                                            $fileSize = Format-FileSize -Bytes $fileInfo.Length
                                            $infoText = "Tamanho: $fileSize"

                                            $ext = $fileInfo.Extension.ToLower()
                                            if ($ext -match "\.(jpg|jpeg|png|gif|bmp|webp|mp4|avi|mkv|mov|wmv)$") {
                                                
                                                try {
                                                    $shellApp = New-Object -ComObject Shell.Application
                                                    $folderInfo = $shellApp.NameSpace($fileInfo.DirectoryName)
                                                    
                                                    if ($folderInfo -ne $null) {
                                                        $shellFile = $folderInfo.ParseName($fileInfo.Name)
                                                        
                                                        if ($shellFile -ne $null) {
                                                            $width = $shellFile.ExtendedProperty("System.Image.HorizontalSize")
                                                            $height = $shellFile.ExtendedProperty("System.Image.VerticalSize")
                                                            
                                                            if ([string]::IsNullOrWhiteSpace($width)) {
                                                                $width = $shellFile.ExtendedProperty("System.Video.FrameWidth")
                                                                $height = $shellFile.ExtendedProperty("System.Video.FrameHeight")
                                                            }
                                                            
                                                            if (-not [string]::IsNullOrWhiteSpace($width) -and -not [string]::IsNullOrWhiteSpace($height)) {
                                                                $infoText = "$width x $height   |   $infoText"
                                                            }
                                                            
                                                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellFile) | Out-Null
                                                        }
                                                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folderInfo) | Out-Null
                                                    }
                                                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellApp) | Out-Null
                                                } catch {}
                                            }
                                            
                                            $bag.PrevStatus.Text = $infoText
                                        } catch {
                                            $bag.PrevStatus.Text = "Erro Metadados: $($_.Exception.Message)"
                                        }
                                    }
                                    # -----------------------------------------------------------

                                    $oldBtn = $bag.PrevForm.Controls | Where-Object { $_.Name -eq "btnDLWV2" }
                                    if ($oldBtn) { $bag.PrevForm.Controls.Remove($oldBtn); $oldBtn.Dispose() }

                                    $bag.PrevRTB.Clear()

                                    if ([string]::IsNullOrWhiteSpace($selItemPath)) {
                                        & $bag.SwitchProf "A"
                                        $bag.PrevLbl.Text = "Selecione um arquivo..."
                                    } elseif (Test-Path -LiteralPath $selItemPath -PathType Container) {
                                        & $bag.SwitchProf "A"
                                        $bag.PrevLbl.Text = "Pasta selecionada.`nNão há preview para pastas."
                                    } else {
                                        $ext = [System.IO.Path]::GetExtension($selItemPath).ToLower()
                                        
                                        # --- A GRANDE MÁGICA: FILTRO INTELIGENTE E AÇÃO ---
                                        $targetProfile = "A"
                                        # Apenas os formatos PESADOS DE VÍDEO acionam o bloqueio de tela (Perfil B)
                                        # Áudio nativo e WebView2 continuam livres usando a memória do Perfil A
                                        if ($ext -match "\.(mp4|avi|wmv|mov|mkv|mpg|mpeg|asf|m4v)$") { $targetProfile = "B" }
                                        & $bag.SwitchProf $targetProfile
                                        # --------------------------------------------------
                                        
                                        if ($ext -match "\.(pdf|webp|webm|svg|avif|apng|ogg|ogv|mht|mhtml)$") {
                                            if ($bag.PrevWV -ne $null) {
                                                try {
                                                    $bag.PrevLbl.Visible = $false
                                                    $bag.PrevWV.Visible = $true
                                                    $bag.PrevWV.BringToFront()
                                                    $bag.PrevWV.Source = New-Object System.Uri($selItemPath)
                                                } catch {
                                                    $bag.PrevLbl.Text = "Erro ao renderizar com o WebView2."
                                                    $bag.PrevLbl.Visible = $true
                                                    $bag.PrevWV.Visible = $false
                                                }
                                            } else {
                                                $bag.PrevLbl.Text = "Formato moderno ($ext) detectado.`n`nPara visualizar este arquivo, o Clone Commander`nprecisa do pacote nativo 'WebView2' (Motor do Edge).`n`nDeseja baixar as DLLs na pasta Config?"
                                                $bag.PrevLbl.Visible = $true
                                                
                                                $btnDl = New-Object System.Windows.Forms.Button
                                                $btnDl.Name = "btnDLWV2"
                                                $btnDl.Text = "Baixar WebView2 (Aprox. 5MB)"
                                                $btnDl.Size = New-Object System.Drawing.Size(200, 35)
                                                $btnDl.Location = New-Object System.Drawing.Point([int](($bag.PrevForm.Width - 200) / 2), [int]($bag.PrevForm.Height / 2 + 60))
                                                $btnDl.BackColor = "#0078D7"
                                                $btnDl.ForeColor = "White"
                                                $btnDl.FlatStyle = "Flat"
                                                $btnDl.Cursor = [System.Windows.Forms.Cursors]::Hand
                                                
                                                # --- DOWNLOAD ASSÍNCRONO BLINDADO ---
                                                $btnDl.Add_Click({
                                                    $this.Text = "Baixando... Aguarde!"
                                                    $this.BackColor = "#555555"
                                                    $this.Enabled = $false
                                                    $bag.PrevForm.Refresh() 
                                                    
                                                    $dlScript = {
                                                        param($configDir, $form, $lbl, $btn)
                                                        try {
                                                            if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Force -Path $configDir | Out-Null }
                                                            
                                                            $nugetUrl = "https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2"
                                                            $tempZip = Join-Path $configDir "webview2_package.zip"
                                                            Invoke-WebRequest -Uri $nugetUrl -OutFile $tempZip -UseBasicParsing
                                                            
                                                            Add-Type -AssemblyName System.IO.Compression.FileSystem
                                                            $zip = [System.IO.Compression.ZipFile]::OpenRead($tempZip)
                                                            
                                                            $core = @($zip.Entries | Where-Object { $_.Name -eq "Microsoft.Web.WebView2.Core.dll" })[0]
                                                            $winforms = @($zip.Entries | Where-Object { $_.Name -eq "Microsoft.Web.WebView2.WinForms.dll" })[0]
                                                            $loader = @($zip.Entries | Where-Object { $_.Name -eq "WebView2Loader.dll" -and $_.FullName -match "win-x64" })[0]
                                                            
                                                            if ($null -eq $core -or $null -eq $winforms -or $null -eq $loader) { throw "Arquivos internos nao encontrados no pacote." }
                                                            
                                                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($core, (Join-Path $configDir "Microsoft.Web.WebView2.Core.dll"), $true)
                                                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($winforms, (Join-Path $configDir "Microsoft.Web.WebView2.WinForms.dll"), $true)
                                                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($loader, (Join-Path $configDir "WebView2Loader.dll"), $true)
                                                            
                                                            $zip.Dispose()
                                                            Remove-Item $tempZip -Force
                                                            
                                                            $form.Invoke([System.Action]{
                                                                $lbl.Text = "Download concluído com sucesso!`n`nFeche e abra o Clone Commander novamente."
                                                                $btn.Visible = $false
                                                            })
                                                        } catch {
                                                            if ($zip -ne $null) { try { $zip.Dispose() } catch {} }
                                                            $errMsg = $_.Exception.Message
                                                            $form.Invoke([System.Action]{
                                                                $lbl.Text = "Erro durante o download/extração.`n`nDetalhe: $errMsg"
                                                                $btn.Text = "Tentar Novamente"
                                                                $btn.BackColor = "#D70000"
                                                                $btn.Enabled = $true
                                                            })
                                                        }
                                                    }
                                                    
                                                    $psDl = [powershell]::Create()
                                                    $psDl.RunspacePool = $global:RunspacePool
                                                    [void]$psDl.AddScript($dlScript)
                                                    
                                                    $baseDir = $global:AppRoot
                                                    $cDir = Join-Path $baseDir "config\WebView2"
                                                    
                                                    [void]$psDl.AddArgument($cDir)
                                                    [void]$psDl.AddArgument($global:SyncHash.Form)
                                                    [void]$psDl.AddArgument($bag.PrevLbl)
                                                    [void]$psDl.AddArgument($this)
                                                    [void]$psDl.BeginInvoke()
                                                }.GetNewClosure())

                                                $bag.PrevForm.Controls.Add($btnDl)
                                                $btnDl.BringToFront()
                                            }
                                        } elseif ($ext -match "\.(gif)$") {
                                            
                                            $bag.PrevLbl.Text = "Carregando GIF animado..."
                                            $bag.PrevLbl.Visible = $true
                                            
                                            try {
                                                # Lê os bytes para a RAM. Isso libera o arquivo original no disco (sem lock), permitindo renomear/excluir!
                                                $bytes = [System.IO.File]::ReadAllBytes($selItemPath)
                                                $ms = New-Object System.IO.MemoryStream(,$bytes)
                                                $gifImg = [System.Drawing.Image]::FromStream($ms)

                                                # O Segredo: Guardamos o MemoryStream na propriedade .Tag para o Garbage Collector não quebrar a animação
                                                $bag.PrevGif.Tag = $ms 
                                                $bag.PrevGif.Image = $gifImg

                                                $bag.PrevLbl.Visible = $false
                                                
                                                # MÁGICA DA CAMADA: Traz o GIF para a frente de tudo
                                                $bag.PrevGif.BringToFront()
                                                $bag.PrevGif.Visible = $true
                                            } catch {
                                                $bag.PrevLbl.Text = "Erro ao carregar a animação do GIF."
                                                $bag.PrevLbl.Visible = $true
                                                $bag.PrevGif.Visible = $false
                                            }

                                        } elseif ($ext -match "\.(jpg|jpeg|png|bmp|ico)$") {
                                            
                                            $bag.PrevLbl.Text = "Carregando WPF..."
                                            $bag.PrevLbl.Visible = $true
                                            
                                            $imgScript = {
                                                param($path, $form, $wpfHost, $lbl)
                                                try {
                                                    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                                                    $bmp.BeginInit()
                                                    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                                                    $bmp.UriSource = New-Object System.Uri($path)
                                                    $bmp.EndInit()
                                                    $bmp.Freeze()
                                                    
                                                    $form.Invoke([System.Action]{
                                                        try {
                                                            $wpfHost.Child.Source = $bmp
                                                            
                                                            # MOTOR TURBO DE QUALIDADE: Força o WPF a renderizar em Alta Qualidade
                                                            [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($wpfHost.Child, [System.Windows.Media.BitmapScalingMode]::HighQuality)
                                                            
                                                            $lbl.Visible = $false
                                                            
                                                            # MÁGICA DA CAMADA: Traz o WPF para a frente de tudo
                                                            $wpfHost.BringToFront()
                                                            $wpfHost.Visible = $true
                                                        } catch {
                                                            $lbl.Text = "Erro ao exibir a imagem no WPF."
                                                        }
                                                    })
                                                } catch {
                                                    $form.Invoke([System.Action]{
                                                        $lbl.Text = "Erro ao ler a imagem nativa."
                                                    })
                                                }
                                            }
                                            
                                            $psImg = [powershell]::Create()
                                            $psImg.RunspacePool = $global:RunspacePool
                                            [void]$psImg.AddScript($imgScript)
                                            [void]$psImg.AddArgument($selItemPath)
                                            [void]$psImg.AddArgument($global:SyncHash.Form)
                                            [void]$psImg.AddArgument($bag.PrevPB)
                                            [void]$psImg.AddArgument($bag.PrevLbl)
                                            [void]$psImg.BeginInvoke()
                                            
                                        } elseif ($ext -match "\.(txt|ps1|bat|csv|json|xml|ini|log|md|js|html|css|py)$") {
                                            
                                            # DELEGA A LEITURA DO TEXTO PARA O SEGUNDO PLANO
                                            $bag.PrevLbl.Text = "Lendo texto..."
                                            $bag.PrevLbl.Visible = $true
                                            
                                            $txtScript = {
                                                param($path, $form, $rtb, $lbl)
                                                try {
                                                    $reader = New-Object System.IO.StreamReader($path)
                                                    $text = ""
                                                    for ($ln = 0; $ln -lt 150 -and -not $reader.EndOfStream; $ln++) { $text += $reader.ReadLine() + "`n" }
                                                    $reader.Close()
                                                    
                                                    $fileLen = (Get-Item $path).Length
                                                    if ($fileLen -gt ($text.Length + 100)) { $text += "`n`n... [ARQUIVO LONGO TRUNCADO PARA PREVIEW] ..." }
                                                    
                                                    $form.Invoke([System.Action]{
                                                        $rtb.Text = $text
                                                        $lbl.Visible = $false
                                                        $rtb.Visible = $true
                                                    })
                                                } catch {
                                                    $form.Invoke([System.Action]{
                                                        $lbl.Text = "Erro ao ler arquivo de texto."
                                                    })
                                                }
                                            }
                                            
                                            $psTxt = [powershell]::Create()
                                            $psTxt.RunspacePool = $global:RunspacePool
                                            [void]$psTxt.AddScript($txtScript)
                                            [void]$psTxt.AddArgument($selItemPath)
                                            [void]$psTxt.AddArgument($global:SyncHash.Form)
                                            [void]$psTxt.AddArgument($bag.PrevRTB)
                                            [void]$psTxt.AddArgument($bag.PrevLbl)
                                            [void]$psTxt.BeginInvoke()
                                            
                                        } elseif ($ext -match "\.(mp4|mp3|avi|wmv|wav|m4a|mov|mkv|flac|aac|mpg|mpeg|wma|asf|mid|midi|m4v)$") {
                                            try {
                                                $bag.PrevLbl.Visible = $false
                                                $bag.PrevWMP.Visible = $true
                                                $bag.PrevWMP.BringToFront()
                                                $bag.PrevWMP.Dock = "None"; $bag.PrevWMP.Dock = "Fill"
                                                
                                                $ocx = $bag.PrevWMP.GetMediaPlayer()
                                                $ocx.uiMode = "full" 
                                                $ocx.windowlessVideo = $true
                                                
                                                $ocx.settings.setMode("loop", $true)
                                                $ocx.settings.autoStart = $true 
                                                $ocx.URL = $selItemPath
                                                
                                                $bag.PrevWMP.Refresh()
                                            } catch {
                                                $bag.PrevLbl.Text = "Erro ao carregar formato de mídia."
                                                $bag.PrevLbl.Visible = $true
                                                $bag.PrevWMP.Visible = $false
                                            }
                                        } else {
                                            $bag.PrevLbl.Text = "Sem preview disponível para $ext"
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                $diskSig = ""
                if ($curPath -match "^([A-Za-z]:)") {
                    $driveLetter = $matches[1]
                    try {
                        $driveInfo = New-Object System.IO.DriveInfo($driveLetter)
                        if ($driveInfo.IsReady) {
                            $volName = $driveInfo.VolumeLabel
                            if ([string]::IsNullOrWhiteSpace($volName)) { $volName = "Disco Local" }
                            $free = Format-FileSize -Bytes $driveInfo.AvailableFreeSpace
                            $total = Format-FileSize -Bytes $driveInfo.TotalSize
                            $diskSig = "$volName ($driveLetter)  |  Livre: $free de $total "
                        }
                    } catch {}
                }
                
                if ($diskSig -ne $bag.LastDiskSig) {
                    $bag.LastDiskSig = $diskSig
                    $bag.DiskLabel.Text = $diskSig
                }
                
                # ==============================================================
                # CORREÇÃO DE LEAK: FAXINA DE COM OBJECTS (ATUALIZADA E BLINDADA)
                # ==============================================================
                if ($null -ne $items) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null }
                if ($null -ne $allItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($allItems) | Out-Null }
                if ($null -ne $selfObj) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selfObj) | Out-Null }
                if ($null -ne $folder) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null }
                if ($null -ne $shellView) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null }
                
            }
        } catch { }
    }.GetNewClosure())
    
    $monitorTimer.Start()

    # ==============================================================
    # MICRO-LEAK CORRIGIDO: Desliga o relógio desta aba assim que o painel for destruído
    # ==============================================================
    $panel.Add_Disposed({
        try {
            if ($monitorTimer) {
                $monitorTimer.Stop()
                $monitorTimer.Dispose()
            }
        } catch {}
    }.GetNewClosure())

    if ($ColumnIndex -eq 0) { $global:LeftTabControl = $tabControl } else { $global:RightTabControl = $tabControl }

    if ($InitialPaths -and $InitialPaths.Count -gt 0) {
        foreach ($path in $InitialPaths) {
            & $AddTabLogic -TargetTabControl $tabControl -Path $path -AddressBox $txtPath -bBack $btnBack -bFwd $btnFwd -BtnView $btnViewMode
        }
        if ($tabControl.TabCount -gt 0) { $tabControl.SelectedIndex = 0 }
    } else {
        & $AddTabLogic -TargetTabControl $tabControl -Path "C:\" -AddressBox $txtPath -bBack $btnBack -bFwd $btnFwd -BtnView $btnViewMode
    }
    
    $tabControl.Tag.BtnAdd = $btnNewTab 
    
    return $tabControl
}

# ==============================================================================
# --- ATALHOS DE TECLADO (RADAR DE MENSAGENS C# E RENOMEIO EM CADEIA) ---
# ==============================================================================
$csharpShortcuts = @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class ShortcutFilter : IMessageFilter {
    // Espião do Windows: Lê qual janela está focada no momento exato do clique
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr GetFocus();
    
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    public Action OnCtrlT;
    public Action OnCtrlShiftT;
    public Action OnGoBack;
    public Action OnGoForward;

    public bool PreFilterMessage(ref Message m) {
        // Intercepta pressionamento de teclas padrão
        if (m.Msg == 0x0100 || m.Msg == 0x0104) { 
            Keys keyData = (Keys)m.WParam.ToInt32() | Control.ModifierKeys;
            
            if (keyData == (Keys.Control | Keys.T)) {
                if (OnCtrlT != null) OnCtrlT();
                return true; 
            }
            if (keyData == (Keys.Control | Keys.Shift | Keys.T)) {
                if (OnCtrlShiftT != null) OnCtrlShiftT();
                return true;
            }

            // --- A MÁGICA DO XYPLORER (RENOMEIO EM CADEIA) ---
            Keys rawKey = (Keys)m.WParam.ToInt32();
            if (rawKey == Keys.Down || rawKey == Keys.Up) {
                IntPtr hFocus = GetFocus();
                if (hFocus != IntPtr.Zero) {
                    StringBuilder sb = new StringBuilder(256);
                    GetClassName(hFocus, sb, 256);
                    
                    // Se a janela piscando for a "Edit" (Caixinha nativa de renomear do Explorer)
                    if (sb.ToString() == "Edit") {
                        if (rawKey == Keys.Down) {
                            SendKeys.Send("{TAB}");
                            return true; // Engole a seta original
                        } else if (rawKey == Keys.Up) {
                            SendKeys.Send("+{TAB}"); // Shift+Tab
                            return true; // Engole a seta original
                        }
                    }
                }
            }
        }
        
        // Teclados com botões Multimídia (Voltar/Avançar)
        if (m.Msg == 0x0319) { 
            int cmd = (int)((m.LParam.ToInt64() >> 16) & 0xFFFF);
            if (cmd == 1) { if (OnGoBack != null) OnGoBack(); return true; }
            if (cmd == 2) { if (OnGoForward != null) OnGoForward(); return true; }
        }
        
        // Mouse com botões laterais (XButton1 e XButton2)
        if (m.Msg == 0x020B) { 
            int xBtn = (int)((m.WParam.ToInt64() >> 16) & 0xFFFF);
            if (xBtn == 1) { if (OnGoBack != null) OnGoBack(); return true; }
            if (xBtn == 2) { if (OnGoForward != null) OnGoForward(); return true; }
        }
        
        return false;
    }
}
'@

if (-not ([System.Management.Automation.PSTypeName]'ShortcutFilter').Type) {
    try { Add-Type -TypeDefinition $csharpShortcuts -ReferencedAssemblies "System.Windows.Forms" -ErrorAction Stop } catch {}
}

function Execute-NewTabShortcut {
    param([bool]$OppositeSide)
    
    if (-not $global:ActiveBrowser) { return }
    
    try {
        $shellView = $global:ActiveBrowser.ActiveXInstance.Document
        
        # Verifica se tem EXATAMENTE UMA pasta selecionada
        if ($shellView) {
            $selItems = $shellView.SelectedItems()
            if ($selItems.Count -eq 1) {
                $item = $selItems.Item(0)
                if ($item.IsFolder) {
                    $targetPath = $item.Path
                    
                    # 1. Descobre em qual lado o seu mouse está clicado agora
                    $isLeft = $false
                    if ($global:LeftTabControl) {
                        foreach ($tab in $global:LeftTabControl.TabPages) {
                            if ($tab.Controls.Count -gt 0 -and $tab.Controls[0] -eq $global:ActiveBrowser) { 
                                $isLeft = $true; break 
                            }
                        }
                    }
                    
                    # 2. Decide o lado alvo
                    $targetTabCtrl = if ($OppositeSide) {
                        if ($isLeft) { $global:RightTabControl } else { $global:LeftTabControl }
                    } else {
                        if ($isLeft) { $global:LeftTabControl } else { $global:RightTabControl }
                    }
                    
                    # 3. Dispara o clique no botão "+" invisivelmente
                    if ($targetTabCtrl -and $targetTabCtrl.Tag -and $targetTabCtrl.Tag.BtnAdd) {
                        $targetTabCtrl.Tag.BtnAdd.PerformClick()
                        
                        Start-Sleep -Milliseconds 50
                        [System.Windows.Forms.Application]::DoEvents()
                        
                        if ($targetTabCtrl.SelectedTab -and $targetTabCtrl.SelectedTab.Controls.Count -gt 0) {
                            $newBrowser = $targetTabCtrl.SelectedTab.Controls[0]
                            $newBrowser.Navigate($targetPath)
                            
                            $shellView.SelectItem($item, 0) # Remove a seleção visual
                        }
                    }
                }
                # --- FAXINA DO ITEM ---
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null
            }
            # --- FAXINA DA COLEÇÃO E DO SHELLVIEW ---
            if ($selItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selItems) | Out-Null }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellView) | Out-Null
        }
    } catch {}
}

# ====================================================================
# --- INIT (CÉREBRO DE INICIALIZAÇÃO E RECOVERY) ---
# ====================================================================

# Liga o Radar à nossa janela principal ANTES de iniciar as abas
$global:ShortcutHandler = New-Object ShortcutFilter
$global:ShortcutHandler.OnCtrlT = [System.Action]{ Execute-NewTabShortcut -OppositeSide $false }
$global:ShortcutHandler.OnCtrlShiftT = [System.Action]{ Execute-NewTabShortcut -OppositeSide $true }

# LIGAÇÃO DOS BOTÕES LATERAIS DO MOUSE E TECLADO
$global:ShortcutHandler.OnGoBack = [System.Action]{ if ($global:ActiveBrowser -and $global:ActiveBrowser.CanGoBack) { $global:ActiveBrowser.GoBack() } }
$global:ShortcutHandler.OnGoForward = [System.Action]{ if ($global:ActiveBrowser -and $global:ActiveBrowser.CanGoForward) { $global:ActiveBrowser.GoForward() } }

[System.Windows.Forms.Application]::AddMessageFilter($global:ShortcutHandler)

# 1. Tenta buscar a última sessão salva silenciosamente
$recoveryData = $null
if (Get-Command "Get-TabSession" -ErrorAction SilentlyContinue) {
    $recoveryData = Get-TabSession
}

# 2. Define os caminhos padrão para a primeira inicialização (Ambos abrem "Meu Computador")
$defaultLeft = @("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") 
$defaultRight = @("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") 

# 3. Decide o que vai abrir (Recovery ou Padrão)
$pathsLeft = if ($recoveryData -and $recoveryData.LeftTabs) { $recoveryData.LeftTabs } else { $defaultLeft }
$pathsRight = if ($recoveryData -and $recoveryData.RightTabs) { $recoveryData.RightTabs } else { $defaultRight }

# 4. Inicializa os painéis (agora enviando uma Lista de caminhos)
$script:leftB  = New-BrowserPane -TableControl $mainTable -ColumnIndex 0 -InitialPaths $pathsLeft
$script:rightB = New-BrowserPane -TableControl $mainTable -ColumnIndex 2 -InitialPaths $pathsRight

# Atualização global segura
if ($global:LeftTabControl.TabPages.Count -gt 0) {
    $global:LeftBrowserRef = $global:LeftTabControl.TabPages[0].Controls[0]
}
if ($global:RightTabControl.TabPages.Count -gt 0) {
    $global:RightBrowserRef = $global:RightTabControl.TabPages[0].Controls[0]
}
$global:ActiveBrowser = $global:LeftBrowserRef

# Garante a última foto das abas, REMOVE o radar para limpar a memória e libera o sistema antes de fechar a janela
$form.Add_FormClosing({ 
    if (Get-Command "Save-TabSession" -ErrorAction SilentlyContinue) { Save-TabSession }
    
    try { [System.Windows.Forms.Application]::RemoveMessageFilter($global:ShortcutHandler) } catch {}
    
    # --- CORREÇÃO DE LEAK (PROCESSO ZUMBI) RESTAURADA ---
    # Desliga o coração do script para o powershell.exe fechar de verdade no Gestor de Tarefas
    try { 
        if ($directionTimer) { 
            $directionTimer.Stop()
            $directionTimer.Dispose() 
        } 
    } catch {}
    
    # Limpeza silenciosa do Mutex
    try {
        if ($null -ne $mutex) {
            $mutex.ReleaseMutex()
            $mutex.Dispose()
        }
    } catch {}
})

# ====================================================================
# CONGELAMENTO VISUAL (Fim): Libera a pintura toda de uma vez só!
# ====================================================================
$form.ResumeLayout($false)
$form.PerformLayout() # Dá um último aviso pro Windows recalcular as larguras invisivelmente

[void]$form.ShowDialog()

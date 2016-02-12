Clear-Host
Stop-Process -name WINWORD
$path = 'C:\'
$files = Get-Childitem $path -Include *.docx,*.doc -Recurse | Where-Object { !($_.psiscontainer) }
$application = New-Object -comobject word.application
$application.visible = $False
$oldlink = '#'
$newlink = '#'

Function Get-StringMatch
{
    Foreach ($file In $files)
    {
        $save = $false
        $document = $application.documents.open($file.fullname,$false,$false)
        $hyperlinks = @($document.Hyperlinks) 
        foreach($hyperlink In $hyperlinks) 
        {
            if($hyperlink.Address -eq $oldlink){
                try{
                    $hyperlink.Address = $newlink
                    $save = $true
                }catch{
                    continue
                }
            }   
        }
        if($save -eq $true){
            try{
                $document.save()
            }
            catch{
                continue
            }
        }
	    $document.close()
    }
    
    $application.quit()
}

Get-StringMatch


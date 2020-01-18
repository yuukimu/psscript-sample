function OverrideMethod ($Document) {
    $doc = $Document | Add-Member -MemberType ScriptMethod -Name "getElementById" -Value {
        param($Id)
        [System.__ComObject].InvokeMember(
            "getElementById",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $this,
            $Id
        ) | ? {$_ -ne [System.DBNull]::Value}
    } -Force -PassThru

    $doc | Add-Member -MemberType ScriptMethod -Name "getElementsByClassName" -Value {
        param($ClassName)
        [System.__ComObject].InvokeMember(
            "getElementsByClassName",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $this,
            $ClassName
        ) | ? {$_ -ne [System.DBNull]::Value}
    } -Force

    $doc | Add-Member -MemberType ScriptMethod -Name "getElementsByTagName" -Value {
        param($TagName)
        [System.__ComObject].InvokeMember(
            "getElementsByTagName",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $this,
            $TagName
        ) | ? {$_ -ne [System.DBNull]::Value}
    } -Force

    $doc | Add-Member -MemberType ScriptMethod -Name "getElementsByName" -Value {
        param($TagName)
        [System.__ComObject].InvokeMember(
            "getElementsByName",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $this,
            $TagName
        ) | ? {$_ -ne [System.DBNull]::Value}
    } -Force

    return $doc
}

function Main($args) {
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Visible = $true
    Google $ie $args
    Qiita $ie
}

function Google($ie, $keyword) {
    $url = "https://google.co.jp/"
    $ie.Navigate($url, 4)
    while ($ie.busy -or $ie.readystate -ne 4)
    {
        Start-Sleep -Milliseconds 100
    }
    Start-Sleep -Milliseconds 100
    $doc = OverrideMethod($ie.Document)
    while ($true) {
        $q = $doc.getElementsByName('q')
        if($q) {break}
    }
    $q[0].value = "藤田茜"
}

function Qiita($ie) {
    $url = "https://qiita.com/"
    $ie.Navigate($url, 4)
    while ($ie.busy -or $ie.readystate -ne 4)
    {
        Start-Sleep -Milliseconds 100
    }
    $doc = OverrideMethod($ie.Document)
    $link = $doc.getElementsByTagName('a') |
    where-object {
        $_.innerText -eq 'ログイン'
    }
    $link.click()
    return
}

Main $args
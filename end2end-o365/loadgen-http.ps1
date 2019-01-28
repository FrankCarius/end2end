while ($true) { ´
	708/(measure-command {`
		$null= Invoke-WebRequest `
			-uri "https://outlook.office365.com/owa/smime/owasmime.msi" `
            -UseBasicParsing
		}`
	).totalseconds`
}


@{
    Severity = @('Error', 'Warning')

    IncludeRules = @(
        'PSAvoidUsingAllowUnencryptedAuthentication',
        'PSAvoidUsingBrokenHashAlgorithms',
        'PSAvoidUsingComputerNameHardcoded',
        'PSAvoidUsingConvertToSecureStringWithPlainText',
        'PSAvoidUsingInvokeExpression',
        'PSAvoidUsingPlainTextForPassword',
        'PSAvoidUsingUsernameAndPasswordParams',
        'PSAvoidUsingWMICmdlet',
        'PSUsePSCredentialType',
        'PSUseShouldProcessForStateChangingFunctions'
    )

    ExcludeRules = @(
        # Add only documented, reviewed exclusions here.
    )
}

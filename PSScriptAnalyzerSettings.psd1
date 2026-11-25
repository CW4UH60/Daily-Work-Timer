@{
    IncludeRules = @(
        'PSAvoidUsingAllowUnencryptedAuthentication',
        'PSAvoidUsingBrokenHashAlgorithms',
        'PSAvoidUsingComputerNameHardcoded',
        'PSAvoidUsingConvertToSecureStringWithPlainText',
        'PSAvoidUsingInvokeExpression',
        'PSAvoidUsingPlainTextForPassword',
        'PSAvoidUsingUsernameAndPasswordParams',
        'PSAvoidUsingWMICmdlet',
        'PSUsePSCredentialType'
    )

    Severity = @(
        'Error',
        'Warning'
    )
}

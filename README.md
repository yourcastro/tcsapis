Public Function GetScoredCardForOilGasExplorationProductionCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasExplorationProductionCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

     Public Function GetScoredCardForOilGasExplorationProductionNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasExplorationProductionNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasMidstreamCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasMidstreamCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

     Public Function GetScoredCardForOilGasMidstreamNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasMidstreamNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasOilFieldServicesCIQ(ByVal filename As String, ByRef ID As Integer,
ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasOilFieldServicesCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasOilFieldServicesNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasOilFieldServicesNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasRefiningMarketingCIQ(ByVal filename As String, ByRef ID As Integer,
ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasRefiningMarketingCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasRefiningMarketingNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasRefiningMarketingNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

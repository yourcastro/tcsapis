Public Function GetScoredCardForCommerialBank(ByVal filename As String, ByRef ID As Integer,
        ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForCommerialBank(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetScoredcardForGenericLargeCorporate(ByVal filename As String, ByRef ID As Integer,
  ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredcardForGenericLargeCorporate(AppSettings.Get(CONFIG_APPNAME),
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

    Public Function GetScoredCardForInsuranceCIQ(ByVal filename As String, ByRef ID As Integer,
       ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForInsuranceCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForInsuranceNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForInsuranceNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetScoredcardForLeaseFinance(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForLeaseFinance(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredcardForMortgageBonds(ByVal filename As String, ByRef ID As Integer,
        ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForMortgageBonds(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForNonBankFinancialCIQ(ByVal filename As String, ByRef ID As Integer,
       ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForNonBankFinancialCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForNonBankFinancialNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForNonBankFinancialNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

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

    Public Function GetScoredCardForUtilities(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForUtilities(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForProjFinance(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForProjFinance(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForPublicFinancNonUS(ByVal filename As String, ByRef ID As Integer,
        ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_IG As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForPublicFinancNonUS(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_IG)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForRealEstateInvestment(ByVal filename As String, ByRef ID As Integer,
ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForRealEstateInvestment(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredcardForSmallMediumEnterprises(ByVal filename As String, ByRef ID As Integer,
   ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredcardForSmallMediumEnterprises(AppSettings.Get(CONFIG_APPNAME),
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

    Public Function GetScoredCardForUniversitySchoolHospital(ByVal filename As String, ByRef ID As Integer,
ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForUniversitySchoolHospital(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function SaveScoreCardData(ByVal fileID As Integer, ByVal filename As String, ByVal userID As String, ByVal xml As String) As String
        Try
            Dim proc As New ClsPDScoreCardProcesor(fileID, filename, userID,
                AppSettings.Get(CONFIG_APPNAME),
                AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                AppSettings.Get(CONFIG_DB_USER),
                AppSettings.Get(CONFIG_DB_PWD),
                AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                CreateSessionId(), xml)
            proc.WriteDB()

            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

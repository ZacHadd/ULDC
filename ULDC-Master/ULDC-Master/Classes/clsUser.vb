' clsUser is a Class with 4 attributes:
'
'       strID ------------- Nexteer users Z-ID.
'       strName ----------- Nexteer user's full name (Last, First Middle).
'       strLicenseLevel --- The Teamcenter license level of the user.
'       strCountry -------- Nexteer user's country.
'
'   Class clsUser also has a function named "PrintAttributes()" which
' will be used to display all of the attributes in a common format.
'
'           (i.e. Pipe Delimited, Tab Delimited, etc)
'
Public Class clsUser

    Private strID As String
    Private strName As String
    Private strLicenseLevel As String
    Private strCountry As String

    ' Default Instance Creation
    '
    Public Sub New()
        strID = "AA0000"
        strName = Nothing
        strLicenseLevel = "None"
        strCountry = Nothing
    End Sub

    ' Named Instance Creation
    Public Sub New(ByVal id As String, ByVal name As String, ByVal license As String, ByVal country As String)
        strID = id
        strName = name
        strLicenseLevel = license
        strCountry = country
    End Sub

    ' Property Getters and Setters
    Public Property ID As String
        Get
            Return (strID)
        End Get
        Set(ByVal value As String)
            strID = value
        End Set
    End Property

    Public Property Name As String
        Get
            Return (strName)
        End Get
        Set(value As String)
            strName = value
        End Set
    End Property

    Public Property LicenseLevel As String
        Get
            Return (strLicenseLevel)
        End Get
        Set(value As String)
            strLicenseLevel = value
        End Set
    End Property

    Public Property Country As String
        Get
            Return (strCountry)
        End Get
        Set(value As String)
            strCountry = value
        End Set
    End Property

    Public Function PrintAttributes()
        ' Here is where the attributes of the Object (id, name, license, country)
        ' will be printed out to the file or where ever it ends up needing to go.

        Return (strID & "|" & strName & "|" & strLicenseLevel & "|" & strCountry & "|")

    End Function

End Class

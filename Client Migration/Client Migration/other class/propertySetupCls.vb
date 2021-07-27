Public Class propertySetupCls
    Property NO As Integer
    Property PROPERTYCODE As String
    Property PROPERTYNAME As String
    Property RESERVATION As Decimal
    Property REMARKS As String
    Property WITHBASEMENT As String
    Property NOS As String
    Property STREET As String
    Property BARANGGAY As String
    Property CITY As String
    Property PROVINCE As String
    Property ZIPCODE As String
    Property COMPID As Integer
    Property MAINID As Integer

    Sub New(no_ As Integer, propertycode_ As String, propertyName_ As String, reservation_ As Decimal, remarks_ As String, withbasement_ As String,
            nos_ As String, street_ As String, barangay_ As String, city_ As String, province_ As String, zipcode_ As String, MAINID As Integer, COMPID As Integer)
        Try
            Me.NO = no_
            Me.PROPERTYCODE = propertycode_
            Me.PROPERTYNAME = propertyName_
            Me.RESERVATION = reservation_
            Me.REMARKS = remarks_
            Me.WITHBASEMENT = withbasement_
            Me.NOS = nos_
            Me.STREET = street_
            Me.BARANGGAY = barangay_
            Me.CITY = city_
            Me.PROVINCE = province_
            Me.ZIPCODE = zipcode_
            Me.MAINID = MAINID
            Me.COMPID = COMPID
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class

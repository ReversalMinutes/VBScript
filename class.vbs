'https://msdn.microsoft.com/en-us/library/4ah5852c(v=vs.84).aspx
Class Customer
   Private m_CustomerName
   Private m_OrderCount

   Private Sub Class_Initialize
      m_CustomerName = ""
      m_OrderCount = 0
   End Sub

   ' CustomerName property.
   Public Property Get CustomerName
      CustomerName = m_CustomerName
   End Property

   Public Property Let CustomerName(custname)
      m_CustomerName = custname
   End Property

   ' OrderCount property (read only).
   Public Property Get OrderCount
      OrderCount = m_OrderCount
   End Property

   ' Methods.
   Public Sub IncreaseOrders(valuetoincrease)
      m_OrderCount = m_OrderCount + valuetoincrease
   End Sub
End Class


Dim c
Set c = New Customer
c.CustomerName = "Fabrikam, Inc."
MsgBox (c.CustomerName)

c.IncreaseOrders(5)
c.IncreaseOrders(3)
MsgBox (c.OrderCount)
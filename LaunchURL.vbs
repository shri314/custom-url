Set App = CreateObject("Shell.Application")
Set Args = WScript.Arguments
If Args.Count > 0 Then
   Loc=InStr( Args.Item(0), "://" )
   If( Loc > 1 ) Then
      Protocol=Left( Args.Item(0), Loc - 1 )
      SrchTerm=Mid( Args.Item(0), InStr( Args.Item(0), "://" ) + 3 )

      REM ########### removing trailing / if it is added implicity

      If ( InStr( SrchTerm, " " ) = 0 AND InStrRev( SrchTerm, "/" ) = Len( SrchTerm ) AND InStr( SrchTerm, "/" ) = Len( SrchTerm ) ) Then
         SrchTerm=Left( SrchTerm, Len(SrchTerm) - 1 )
      End If

      REM ########### encode special chars

      SrchTerm=Replace( SrchTerm, "+", "%2B" )
      SrchTerm=Replace( SrchTerm, " ", "+" )
      SrchTerm=Replace( SrchTerm, "/", "%2F" )
      SrchTerm=Replace( SrchTerm, "\", "%5C" )

      REM ########### dispatch

      If( Protocol = "google" ) Then
         R = App.ShellExecute("http://www.google.com/search?q=" & SrchTerm)
      End If
      If( Protocol = "images" ) Then
         R = App.ShellExecute("http://www.google.com/search?tbm=isch&q=" & SrchTerm)
      End If
      If( Protocol = "bing" ) Then
         R = App.ShellExecute("http://www.bing.com/search?q=" & SrchTerm)
      End If
      If( Protocol = "stack" ) Then
         R = App.ShellExecute("http://stackoverflow.com/search?q=" & SrchTerm)
      End If
   End If
End If


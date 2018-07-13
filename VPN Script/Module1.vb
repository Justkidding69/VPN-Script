Imports DotRas

Module Module1

    Sub Main()
        If Not My.Application.CommandLineArgs.Count = 3 Then 'delete the if else statement if you change the strings below'
            ShowUsage()
        Else 'Change My.Application.CommandLineArgs(0-2) to a string if you want to use this directly'
            Dim VpnName As String = My.Application.CommandLineArgs(0)
            Dim Destination As String = My.Application.CommandLineArgs(1)
            Dim PresharedKey As String = My.Application.CommandLineArgs(2)

            Try
                Dim PhoneBook As New RasPhoneBook
                PhoneBook.Open()
                Dim VpnEntry As RasEntry = RasEntry.CreateVpnEntry(VpnName, Destination, DotRas.RasVpnStrategy.L2tpOnly,
                                                                   DotRas.RasDevice.Create(VpnName, DotRas.RasDeviceType.Vpn))
                'Check what VPN options you want (not all are here)'
                VpnEntry.Options.UsePreSharedKey = True
                VpnEntry.Options.RequirePap = True
                VpnEntry.Options.RequireChap = False
                VpnEntry.Options.RequireMSChap2 = True
                VpnEntry.Options.RequireEncryptedPassword = False
                VpnEntry.Options.UseLogOnCredentials = False
                PhoneBook.Entries.Add(VpnEntry)
                VpnEntry.UpdateCredentials(RasPreSharedKey.Client, PresharedKey)
                Console.WriteLine("VPN connection created successfully")
            Catch ex As Exception
                Console.WriteLine("ERROR: " & ex.Message & vbNewLine)
                Environment.Exit(999)
            End Try
        End If
    End Sub

    Private Sub ShowUsage()
        Console.WriteLine("Invalid number of arguments specified." & vbNewLine & vbNewLine &
                          "Usage: VpnScript.exe [VPN Name] [Destination] [Preshared Key]" & vbNewLine & vbNewLine &
                          "EXAMPLE: VpnScript.exe ""New VPN"" vpn.mycompany.com SomePassword" & vbNewLine)
        Console.ReadKey()
    End Sub

End Module

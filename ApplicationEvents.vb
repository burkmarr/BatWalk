Namespace My

    Partial Friend Class MyApplication

        Protected Overrides Function OnInitialize(ByVal commandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String)) As Boolean
            ' Set the display time to 4000 milliseconds (4 seconds). 
            Me.MinimumSplashScreenDisplayTime = 4000
            Return MyBase.OnInitialize(commandLineArgs)
        End Function

    End Class
End Namespace
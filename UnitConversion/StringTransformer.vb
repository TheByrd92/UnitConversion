Imports System.Text.RegularExpressions
''' <summary>
''' Transforms string in predefined manners with each individual function.
''' </summary>
''' <remarks>Author Austin and Ian.</remarks>
Public Class StringTransformer
    ''' <summary>
    ''' Replaces any match to this regular expression <code>\s?\(?(OPP(OSITE)?\.? HAND|AS SHOWN)\)?</code> with nothing.
    ''' </summary>
    ''' <param name="PieceDescription">The string to compare and replace.</param>
    ''' <returns>The string that has the regular expression removed or if no match is found returns the same string passed to it.</returns>
    ''' <remarks>We should find where this is called in the program.</remarks>
    Public Function RemoveHand(ByVal PieceDescription As String) As String
        PieceDescription = PieceDescription.Trim()
        Dim xprs = New Regex("\s?\(?(OPP(OSITE)?\.? HAND|AS SHOWN)\)?")
        Return If(xprs.Match(PieceDescription).Success, PieceDescription.Replace(xprs.Match(PieceDescription).ToString(), ""), PieceDescription)
    End Function
End Class

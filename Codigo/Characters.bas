Attribute VB_Name = "Characters"
'**************************************************************
' Characters.bas - library of functions to manipulate characters.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

''
' Value representing invalid indexes.
Public Const INVALID_INDEX As Integer = 0

''
' Retrieves the UserList index of the user with the give char index.
'
' @param    CharIndex   The char index being used by the user to be retrieved.
' @return   The index of the user with the char placed in CharIndex or INVALID_INDEX if it's not a user or valid char index.
' @see      INVALID_INDEX

Public Function CharIndexToUserIndex(ByVal CharIndex As Integer) As Integer
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Takes a CharIndex and transforms it into a UserIndex. Returns INVALID_INDEX in case of error.
'***************************************************
    CharIndexToUserIndex = CharList(CharIndex)
    
    If CharIndexToUserIndex < 1 Or CharIndexToUserIndex > MaxUsers Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
    
    If UserList(CharIndexToUserIndex).Char.CharIndex <> CharIndex Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
End Function



Public Function PuedeRecuperar(ByVal UserIndex As Integer, ByVal UserName As String, ByVal email As String, ByVal clave As String) As Boolean
  'Función para ver si el char puede RECUPERAR el personaje.
 
  '//Variables
    Dim Data_mail As String: Data_mail = UCase$(GetVar(CharPath & UserName & ".chr", "CONTACTO", "Email"))
    Dim Data_Pin As String: Data_Pin = UCase$(GetVar(CharPath & UserName & ".chr", "CONTACTO", "Clave"))
   
    With UserList(UserIndex)
   
        If Not PersonajeExiste(UserName) Then
            Call WriteErrorMsg(UserIndex, "El personaje no existe.")
            PuedeRecuperar = False
            Exit Function
        End If
       
        If UCase$(email) <> Data_mail Then
            Call WriteErrorMsg(UserIndex, "El mail ingresado no es correcto.")
            PuedeRecuperar = False
            Exit Function
        End If
       
        If UCase$(clave) <> Data_Pin Then
            Call WriteErrorMsg(UserIndex, "La clave/pin no pertenece al personaje.")
            PuedeRecuperar = True
            Exit Function
        End If
 
        PuedeRecuperar = True
    End With
End Function
Public Function PuedeBorrar(ByVal UserIndex As Integer, ByVal UserName As String, ByVal email As String, ByVal clave As String, ByVal Passwd As String) As Boolean
'Función para ver si puede BORRAR el personaje
    
    'SHA256
    Dim oSHA256 As CSHA256
    Set oSHA256 = New CSHA256
    
    'Datos del personaje
    Dim Data_mail As String: Data_mail = UCase$(GetVar(CharPath & UserName & ".chr", "CONTACTO", "Email"))
    Dim Data_Pin As String: Data_Pin = UCase$(GetVar(CharPath & UserName & ".chr", "CONTACTO", "Clave"))
    Dim Data_Passwd As String: Data_Passwd = GetVar(CharPath & UserName & ".chr", "INIT", "Password")
    Dim Salt As String: Salt = GetVar(CharPath & UserName & ".chr", "INIT", "Salt") ' Obtenemos la Salt
    With UserList(UserIndex)
   
        If Not PersonajeExiste(UserName) Then
            Call WriteErrorMsg(UserIndex, "El personaje no existe.")
            PuedeBorrar = False
            Exit Function
        End If
       
        If UCase$(email) <> Data_mail Then
            Call WriteErrorMsg(UserIndex, "El mail ingresado no es correcto.")
            PuedeBorrar = False
            Exit Function
        End If
       
        If UCase$(clave) <> Data_Pin Then
            Call WriteErrorMsg(UserIndex, "La clave pin no pertenece al personaje.")
            PuedeBorrar = False
            Exit Function
        End If
        
        If Not oSHA256.SHA256(Passwd & Salt) = Data_Passwd Then
            Call WriteErrorMsg(UserIndex, "El password no pertenece al personaje.")
            PuedeBorrar = False
            Exit Function
        End If
       
        PuedeBorrar = True
    End With
End Function

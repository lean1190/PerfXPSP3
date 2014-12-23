VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PerfXPSP3 ::."
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnborrar 
      Caption         =   "ST"
      Height          =   375
      Left            =   3540
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Seleccionar/Deseleccionar todos los items en ""Borrar"""
      Top             =   6260
      Width           =   375
   End
   Begin VB.CommandButton btntweaks 
      Caption         =   "ST"
      Height          =   375
      Left            =   11700
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Seleccionar/Deseleccionar todos los items en ""Tweaks"""
      Top             =   3250
      Width           =   375
   End
   Begin VB.CommandButton btndes 
      Caption         =   "ST"
      Height          =   375
      Left            =   11700
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Seleccionar/Deseleccionar todos los items en ""Deshabilitar"""
      Top             =   260
      Width           =   375
   End
   Begin VB.CommandButton btnsalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Termina el programa"
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton btnnone 
      Caption         =   "Seleccionar &Ninguna"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   "Deselecciona todas las opciones"
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton btnall 
      Caption         =   "Seleccionar &Todas"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Selecciona todas las opciones"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Frame frmtweaks 
      Caption         =   "Tweaks"
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   12015
      Begin VB.CheckBox chk 
         Caption         =   "Quitar número de mensajes no leídos"
         Height          =   375
         Index           =   37
         Left            =   7920
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Quita el mensaje ""Mensajes no leídos x"" de la pantalla de bienvenida"
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "No mostrar errores en el inicio"
         Height          =   375
         Index           =   35
         Left            =   7920
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Evita los mensajes de error al inicio del SO (por ejemplo un error como ""El equipo se ha recuperado de un error grave"")"
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "Dar mas prioridad a la ventana activa"
         Height          =   375
         Index           =   36
         Left            =   7920
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Prioriza la utilización de memoria en la ventana actual"
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "Borrar carpeta Documentos Compartidos"
         Height          =   375
         Index           =   39
         Left            =   7920
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Quita de Mi PC la carpeta Documentos Compartidos"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CheckBox chk 
         Caption         =   "Borrar barra de idioma de la barra de tareas"
         Height          =   375
         Index           =   40
         Left            =   7920
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Borra la barra de idiomas de abajo a la derecha,a la izquierda del reloj"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Borrar documentos recientes al cerrar sesion"
         Height          =   375
         Index           =   38
         Left            =   7920
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Borra la lista de documentos compartidos del menú inicio"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Auto reiniciar escritorio en caso de error"
         Height          =   375
         Index           =   29
         Left            =   3840
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Reinicia el proceso ""explorer.exe"" en caso de cerrarse por algún error"
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox chk 
         Caption         =   "Abrir carpetas en procesos separados"
         Height          =   375
         Index           =   28
         Left            =   3840
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Mejora la estabilidad del sistema al abrir las carpetas separadas del entorno"
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "Abrir aplicaciones de 16-bits en procesos separados"
         Height          =   375
         Index           =   34
         Left            =   3840
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Mejora la estabilidad del sistema al abrir las aplicaciones de 16 bits  separados del entorno"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.CheckBox chk 
         Caption         =   "Agregar ""Abrir en DOS"" en el menu contextual"
         Height          =   375
         Index           =   32
         Left            =   3840
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Agrega la opción ""Abrir en DOS"" cuando se hace click derecho sobre un archivo o carpeta"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CheckBox chk 
         Caption         =   "Mostrar en el escritorio la versión del sistema"
         Height          =   375
         Index           =   41
         Left            =   7920
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Muestra en el escritorio abajo a la derecha la versión y compilación de Windows"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Quitar globos de información al lado del reloj"
         Height          =   375
         Index           =   30
         Left            =   3840
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Elimina los ""globos"" amarillos de información de abajo a la derecha (por ejemplo el de ""Es seguro retirar el hardware"")"
         Top             =   960
         Width           =   3495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Borrar flechitas de accesos directos"
         Height          =   375
         Index           =   24
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Elimina las flechitas de accesos directos"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CheckBox chk 
         Caption         =   "Suprimir la animación de las ventanas"
         Height          =   375
         Index           =   26
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Elimina la animación de las ventanas al minimizar y maximizar"
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "Optimizar el disco duro durante el apagado"
         Height          =   375
         Index           =   27
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Optimiza el disco duro cuando se apaga el sistema"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         Caption         =   "Descargar las DLL no utilizadas"
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Fuerza la descarga de memoria de las DLL no utilizadas"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "Aumentar límite de conexiones simultáneas a 100"
         Height          =   375
         Index           =   33
         Left            =   3840
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta el límite de conexiones simultáneas de 2 a 100"
         Top             =   2040
         Width           =   3855
      End
      Begin VB.CheckBox chk 
         Caption         =   "Mejorar administración de memoria"
         Height          =   375
         Index           =   23
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Mejora el tratamiento de las aplicaciones en memoria"
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox chk 
         Caption         =   "Acelerar acceso en volúmenes NTFS"
         Height          =   375
         Index           =   25
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Mejora el acceso a directorios en volúmenes NTFS"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "Acelerar apagado de Windows"
         Height          =   375
         Index           =   21
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Fuerza las aplicaciones y servicios a cerrar más velozmente"
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "Reducir tiempo de despliegue del menu inicio"
         Height          =   375
         Index           =   31
         Left            =   3840
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Minimiza la animación del menú inicio para desplegarlo más rápidamente"
         Top             =   1320
         Width           =   3495
      End
   End
   Begin VB.CommandButton btnaplicar 
      Caption         =   "&Aplicar"
      Height          =   1358
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "APLICA LOS CAMBIOS AHORA!!! ^^"
      Top             =   6217
      Width           =   1455
   End
   Begin VB.Frame frmdesh 
      Caption         =   "Deshabilitar"
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   12015
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar DISKPERF"
         Height          =   375
         Index           =   18
         Left            =   7920
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita Diskperf (colecciona datos físicos de las unidades de disco y transforma los valores en contadores de rendimiento)"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar rastreo de acceso directo roto"
         Height          =   375
         Index           =   19
         Left            =   7920
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   $"Form1.frx":3AFA
         Top             =   2040
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar servicio web de asociación de archivos"
         Height          =   375
         Index           =   20
         Left            =   7920
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita la opción de buscar un programa asociado con la extensión del archivo desconocido"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar aviso de poco espacio en disco"
         Height          =   375
         Index           =   15
         Left            =   7920
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita el aviso de poco espacio en disco pidiéndonos que borremos información"
         Top             =   600
         Width           =   3495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar StickyKeys"
         Height          =   375
         Index           =   17
         Left            =   7920
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita las ""teclas pegajosas"" (cuando tocas por cierto tiempo la tecla ""Shift"",por ejemplo)"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Servicios innecesarios"
         Height          =   375
         Index           =   16
         Left            =   7920
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   960
         Width           =   2895
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar asistente de búsqueda"
         Height          =   375
         Index           =   13
         Left            =   3840
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita el ""perro"" de la opcion ""Buscar"" y da un entorno más limpio"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar seguimiento de usuario"
         Height          =   375
         Index           =   12
         Left            =   3840
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   $"Form1.frx":3B92
         Top             =   2040
         Width           =   2895
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar asistente de limpieza de escritorio"
         Height          =   375
         Index           =   14
         Left            =   7920
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita el asistente de limpieza de escritorio que se ejecuta cada cierto tiempo automáticamente"
         Top             =   240
         Width           =   3615
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Suspender"
         Height          =   375
         Index           =   11
         Left            =   3840
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita la opcion Suspender/Hibernar"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Reinicio Automático"
         Height          =   375
         Index           =   10
         Left            =   3840
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita el reinicio automático del sistema luego de un error"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Restaurar Sistema"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   $"Form1.frx":3C4E
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Reporte de Errores de Windows"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita el envio de informes de error de Windows (cuando ocurre un error y da la opcion ""Enviar ""o ""No enviar"")"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Reporte de errores de IE"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   $"Form1.frx":3CDD
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar opciones de grabado de Windows"
         Height          =   375
         Index           =   9
         Left            =   3840
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita el grabado de Windows"
         Top             =   960
         Width           =   3615
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Local Group Policy Objects"
         Height          =   375
         Index           =   7
         Left            =   3840
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Conjunto de políticas del sistema (que solo ocupan memoria)"
         Top             =   240
         Width           =   3135
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar DrWatson"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   $"Form1.frx":3D64
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   "Dehabilitar Autorun ( CD + USB )"
         Height          =   375
         Index           =   8
         Left            =   3840
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilitar la opcion de ""autoejecucion"" en lectoras y usb"
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Actualizaciones Automáticas"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita las Actualizaciones Automaticas de Windows"
         Top             =   960
         Width           =   3255
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Alertas de Seguridad"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita las alertas del Centro de Seguridad de Windows (las del centro de seguridad)"
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox chk 
         Caption         =   "Deshabilitar Firewall"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Deshabilita el Firewall de Windows (conseguite un firewall de verdad!)"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frmborrar 
      Caption         =   "Borrar"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   3855
      Begin VB.CheckBox chk 
         Caption         =   "Borrar Windows Messenger (Messenger 5)"
         Height          =   375
         Index           =   44
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Elimina Messenger 5 de una buena vez!!"
         Top             =   960
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         Caption         =   "Borrar Screensavers de Windows"
         Height          =   375
         Index           =   43
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Borra los screensavers básicos de Windows (instalate uno bueno!)"
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox chk 
         Caption         =   "Borrar DLL Cache"
         Height          =   375
         Index           =   42
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Borra los archivos de resguardo de Windows"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   9900
      Picture         =   "Form1.frx":3DF1
      ToolTipText     =   "ABOUT"
      Top             =   7080
      Width           =   2310
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Procesar() 'Procesa las opciones
  Dim i As Integer
    For i = 0 To 44 Step 1
      If (chk(i).Value = 1) Then
        Select Case i
          Case 0
            Deshabilitar_Firewall
          Case 1
            Deshabilitar_Alertas_de_Seguridad
          Case 2
            Deshabilitar_AutoUpdates
          Case 3
            Deshabilitar_Errores_IE
          Case 4
            Deshabilitar_Errores_WIN
          Case 5
            Deshabilitar_DrWatson
          Case 6
            Deshabilitar_Restaurar_Sistema
          Case 7
            Deshabilitar_LGPO
          Case 8
            Deshabilitar_AutoRUN
          Case 9
            Deshabilitar_Grabado
          Case 10
            Deshabilitar_AutoReinicio
          Case 11
            Deshabilitar_Suspender
          Case 12
            Deshabilitar_UserTracking
          Case 13
            Deshabilitar_Asistente_Busqueda
          Case 14
            Deshabilitar_Limpieza_Escritorio
          Case 15
            Deshabilitar_Poco_Espacio
          Case 16
            Deshabilitar_InServs
          Case 17
            Deshabilitar_StickyK
          Case 18
            Deshabilitar_Diskperf
          Case 19
            Deshabilitar_Rastreo
          Case 20
            Deshabilitar_AsociacionWeb
          Case 21
            Acelerar_Apagado
          Case 22
            Descargar_DLL
          Case 23
            Mejorar_Memoria
          Case 24
            Accesos_Directos
          Case 25
            Acceso_NTFS
          Case 26
            Supr_Animacion
          Case 27
            Optimizar_HD
          Case 28
            Carpetas_ProcSep
          Case 29
            AutoReinicio_Desk
          Case 30
            Quitar_Globos
          Case 31
            Reducir_Tiempo_Inicio
          Case 32
            Abrir_DOS
          Case 33
            Aumentar_Conexiones100
          Case 34
            Abrir16bits_ProcSep
          Case 35
            NoErrores_Inicio
          Case 36
            Mas_Prioridad
          Case 37
            Quitar_NumMsjs
          Case 38
            Borrar_ArchRecientes
          Case 39
            Borrar_Documentos_Compartidos
          Case 40
            Borrar_Barra_Idiomas
          Case 41
            Mostrar_Version_Escritorio
          Case 42
            Borrar_DLLCache
          Case 43
            Borrar_Screens
          Case 44
            Borrar_Messenger
        End Select
    End If
  Next
End Sub
Public Function Check(ByRef LimiteInf, LimiteSup As Integer) As Boolean 'Checkea la seleccion de los boxes
  Dim ok As Boolean
  ok = True
  Do While (LimiteSup > LimiteInf - 1) And (ok)
    If (chk(LimiteSup).Value = 0) Then
      ok = False
    End If
    LimiteSup = LimiteSup - 1
  Loop
  Check = ok
End Function
Private Sub btndes_Click() 'Boton Deshabilitar
  If (Check(0, 20)) Then
    Deseleccionar 0, 20
  Else
    Seleccionar 0, 20
  End If
End Sub
Private Sub btntweaks_Click() 'Boton Tweaks
  If (Check(21, 41)) Then
    Deseleccionar 21, 41
  Else
    Seleccionar 21, 41
  End If
End Sub
Private Sub btnborrar_Click() 'Boton Borrar
  If (Check(42, 44)) Then
    Deseleccionar 42, 44
  Else
    Seleccionar 42, 44
  End If
End Sub
Private Sub Form_Load() 'Selecciona todas al principio
  Seleccionar 0, 44
End Sub
Private Sub btnall_Click() 'Boton que selecciona todos los boxes
  Seleccionar 0, 44
End Sub
Private Sub btnnone_Click() 'Boton que deselecciona todos los boxes
  Deseleccionar 0, 44
End Sub
Private Sub btnsalir_Click() 'Boton para salir
  Unload Me
  Set Form1 = Nothing
  End
End Sub
Private Sub btnaplicar_Click() 'Boton aplicar para procesar los cambios
  Dim Respuesta As Long
  Respuesta = MsgBox("==> Este programa modifica entradas en el registro de Windows" & vbCrLf & _
  "==> Si bien no perjudicará en nada al SO es recomendable" & vbCrLf & "==> hacer un BackUp (copia de seguridad) del registro" & vbCrLf & _
  "" & vbCrLf & "==> ¿Comenzar? <==", vbInformation + vbOKCancel, "Antes de comenzar...")
  If (Respuesta = vbOK) Then
    Form1.Enabled = False
    Procesar
    MsgBox "Todos los cambios han sido aplicados satisfactoriamente" _
    & vbCrLf & "" & vbCrLf & "Algunos de estos requieren reiniciar la PC para que se efectuen" & vbCrLf & _
    "" & vbCrLf & "Gracias por usar mi programa! ^^", vbInformation + vbOKOnly, "Listo!"
    Form1.Enabled = True
  ElseIf (Respuesta = vbCancel) Then
    Exit Sub
  End If
End Sub
'----------------------------------DESHABILITAR----------------------------
Public Sub Deshabilitar_Firewall()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\EnableFirewall", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Setup\InterfacesUnfirewalledAtUpdate\All", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\EnableFirewall", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\ControlSet001\Services\SharedAccess\Setup\InterfacesUnfirewalledAtUpdate\All", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\ControlSet002\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\EnableFirewall", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\ControlSet002\Services\SharedAccess\Setup\InterfacesUnfirewalledAtUpdate\All", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\Software\Policies\Microsoft\WindowsFirewall\DomainProfile\EnableFirewall", 0, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Alertas_de_Seguridad()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Security Center\FirstRunDisabled", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Security Center\FirewallDisableNotify", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Security Center\UpdatesDisableNotify", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Security Center\AntiVirusDisableNotify", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Security Center\AntiVirusOverride", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Security Center\FirewallOverride", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\wscsvc\Start", 4, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_AutoUpdates()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\ConfigVer", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\AUOptions", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\wuauserv\Start", 4, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Errores_IE()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Internet Explorer\Main\IEWatsonEnabled", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Internet Explorer\Main\IEWatsonDisabled", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Errores_WIN()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\PCHealth\ErrorReporting\DoReport", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\ERSvc\Start", 4, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_DrWatson()
  On Error Resume Next
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug\"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Restaurar_Sistema()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore\DisableConfig", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore\DisableSR", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\srservice\Start", 4, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_LGPO()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\System\DisableGPO", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_AutoRUN()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Cdrom\AutoRun", 0, "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutoRun", 255, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Grabado()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoCDBurning", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_AutoReinicio()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\AutoReboot", 0, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Suspender()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\ACPI\Parameters\Attributes", 112, "REG_DWORD"
  Set WshShell = Nothing
  Shell "cmd.exe /c REG ADD HKLM\SYSTEM\CurrentControlSet\Services\ACPI\Parameters /v AMLIMaxCTObjs /t REG_BINARY /d 04000000 /f", vbHide
  Shell "cmd.exe /c REG ADD HKLM\SYSTEM\CurrentControlSet\Services\ACPI\Parameters\WakeUp /v FixedEventMask /t REG_BINARY /d 2005 /f", vbHide
  Shell "cmd.exe /c REG ADD HKLM\SYSTEM\CurrentControlSet\Services\ACPI\Parameters\WakeUp /v FixedEventStatus /t REG_BINARY /d 0084 /f", vbHide
  Shell "cmd.exe /c REG ADD HKLM\SYSTEM\CurrentControlSet\Services\ACPI\Parameters\WakeUp /v GenericEventMask /t REG_BINARY /d 18500010 /f", vbHide
  Shell "cmd.exe /c REG ADD HKLM\SYSTEM\CurrentControlSet\Services\ACPI\Parameters\WakeUp /v GenericEventStatus /t REG_BINARY /d 1000ff00 /f", vbHide
End Sub
Public Sub Deshabilitar_UserTracking()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoInstrumentation", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Asistente_Busqueda()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState\Use Search Asst", "no", "REG_SZ"
  WshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\Use Search Asst", "no", "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Limpieza_Escritorio()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\CleanupWiz\NoRun", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Poco_Espacio()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLowDiskSpace", 1, "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLowDiskSpaceChecks", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_InServs()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\WmiApSrv\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\RDSessMgr\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\LmHosts\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\helpsvc\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\NetDDEdsdm\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\RemoteAccess\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\mnmsrvc\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\seclogon\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\RpcLocator\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Messenger\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\NtLmSsp\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\RSVP\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\RemoteRegistry\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\SysmonLog\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\CiSvc\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\SSDPSRV\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\BITS\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\TermService\Start", 4, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\TlntSvr\Start", 4, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_StickyK()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Control Panel\Accessibility\StickyKeys\Flags", "506", "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_Diskperf()
  Shell "cmd.exe /c diskperf -n", vbHide
End Sub
Public Sub Deshabilitar_Rastreo()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoResolveTrack", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Deshabilitar_AsociacionWeb()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system\NoInternetOpenWith", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub

'----------------------------------TWEAKS-----------------------------------

Public Sub Acelerar_Apagado()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Control Panel\Desktop\ForegroundLockTimeout", 0, "REG_DWORD"
  WshShell.RegWrite "HKCU\Control Panel\Desktop\AutoEndTasks", "1", "REG_SZ"
  WshShell.RegWrite "HKCU\Control Panel\Desktop\WaitToKillAppTimeout", "1000", "REG_SZ"
  WshShell.RegWrite "HKCU\Control Panel\Desktop\HungAppTimeout", "1000", "REG_SZ"
  WshShell.RegWrite "HKEY_USERS\.DEFAULT\Control Panel\Desktop\ForegroundLockTimeout", 0, "REG_DWORD"
  WshShell.RegWrite "HKEY_USERS\.DEFAULT\Control Panel\Desktop\AutoEndTasks", "1", "REG_SZ"
  WshShell.RegWrite "HKEY_USERS\.DEFAULT\Control Panel\Desktop\WaitToKillAppTimeout", "1000", "REG_SZ"
  WshShell.RegWrite "HKEY_USERS\.DEFAULT\Control Panel\Desktop\HungAppTimeout", "1000", "REG_SZ"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\WaitToKillAppTimeout", "1000", "REG_SZ"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\WaitToKillServiceTimeout", "1000", "REG_SZ"
  WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Reliability\ShutdownReasonOn", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Reliability\ShutdownReasonUI", 0, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\PowerdownAfterShutdown", "1", "REG_SZ"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\ShutdownWithoutLogon", "1", "REG_SZ"
  WshShell.RegWrite "HKCU\Control Panel\Desktop\PowerOffActive", "1", "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub Descargar_DLL()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AlwaysUnloadDLL", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Mejorar_Memoria()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\DisablePagingExecutive", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\LargeSystemCache", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Accesos_Directos()
  On Error Resume Next
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegDelete "HKCR\lnkfile\IsShortcut"
  WshShell.RegDelete "HKCR\piffile\IsShortcut"
  Shell "cmd.exe /c REG ADD HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer /v link /t REG_BINARY /d 00000000 /f", vbHide
  Set WshShell = Nothing
End Sub
Public Sub Acceso_NTFS()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem\NtfsDisable8dot3NameCreation", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem\NtfsDisableLastAccessUpdate", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Supr_Animacion()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Control Panel\Desktop\WindowMetrics\MinAnimate", "0", "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub Optimizar_HD()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable", "Y", "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub Carpetas_ProcSep()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\SeparateProcess", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub AutoReinicio_Desk()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoRestartShell", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Quitar_Globos()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\EnableBalloonTips", 0, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Reducir_Tiempo_Inicio()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Control Panel\Desktop\MenuShowDelay", "100", "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub Abrir_DOS()
  Dim Cadena As String
  Cadena = "C:\Windows\System32\cmd /k cd " & Chr(34) & Chr(37) & "1" & Chr(34)
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCR\Directory\shell\DosHere\", "Abrir en DOS", "REG_SZ"
  WshShell.RegWrite "HKCR\Directory\shell\DosHere\Command\", Cadena, "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub Aumentar_Conexiones100()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPer1_0Server", 100, "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPerServer", 100, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Abrir16bits_ProcSep()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\WOW\DefaultSeparateVDM", "yes", "REG_SZ"
  Set WshShell = Nothing
End Sub
Public Sub NoErrores_Inicio()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows\NoPopUpsOnBoot", 1, "REG_DWORD"
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows\Error Mode", 2, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Mas_Prioridad()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\PriorityControl\Win32PrioritySeparation", 38, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Quitar_NumMsjs()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\UnreadMail\MessageExpiryDays", 0, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Borrar_ArchRecientes()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\ClearRecentDocsOnExit", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
Public Sub Borrar_Documentos_Compartidos()
  On Error Resume Next
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\DelegateFolders\{59031a47-3f72-44a7-89c5-5595fe6b30ee}\"
  Set WshShell = Nothing
End Sub
Public Sub Borrar_Barra_Idiomas()
  On Error Resume Next
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegDelete "HKEY_CLASSES_ROOT\CLSID\{540D8A8B-1C3F-4E32-8132-530F6A502090}\"
  Set WshShell = Nothing
End Sub
Public Sub Mostrar_Version_Escritorio()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKCU\Control Panel\Desktop\PaintDesktopVersion", 1, "REG_DWORD"
  Set WshShell = Nothing
End Sub
'---------------------------------BORRAR-------------------------------------
Public Sub Borrar_DLLCache()
  Dim WshShell
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SfcQuota", 0, "REG_DWORD"
  Set WshShell = Nothing
  Shell "cmd.exe /c del /f /s /q %systemroot%\system32\dllcache", vbHide
End Sub
Public Sub Borrar_Screens()
  Shell "cmd.exe /c del /f /s /q %systemroot%\system32\*.scr", vbHide
End Sub
Public Sub Borrar_Messenger()
  Shell "cmd.exe /c RunDll32 advpack.dll,LaunchINFSection %windir%\INF\msmsgs.inf,BLC.Remove"
End Sub

Private Sub Image1_Click() 'Creditos
  MsgBox "© Copyright 2010 - 4ever FREEWARE" & vbCrLf & ".::PerfXPSP3 (engrana tu WindowsXP).::" _
  & vbCrLf & "" & vbCrLf & "Gracias por usar mi programa <==", vbInformation + vbOKOnly, "About Box"
End Sub

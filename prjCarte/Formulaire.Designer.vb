<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Formulaire
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label_resultat = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.resultat = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.last_numero = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(75, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Veuillez inserer le numéro"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(78, 103)
        Me.TextBox1.MaxLength = 12
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(267, 20)
        Me.TextBox1.TabIndex = 1
        '
        'Label_resultat
        '
        Me.Label_resultat.AutoSize = True
        Me.Label_resultat.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_resultat.Location = New System.Drawing.Point(75, 249)
        Me.Label_resultat.Name = "Label_resultat"
        Me.Label_resultat.Size = New System.Drawing.Size(54, 13)
        Me.Label_resultat.TabIndex = 2
        Me.Label_resultat.Text = "Resultat"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(130, 140)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(156, 49)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Génerer"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'resultat
        '
        Me.resultat.AutoSize = True
        Me.resultat.Location = New System.Drawing.Point(75, 280)
        Me.resultat.Name = "resultat"
        Me.resultat.Size = New System.Drawing.Size(137, 13)
        Me.resultat.TabIndex = 4
        Me.resultat.Text = "Initials:    Naissance:    Sex:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(73, 377)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Derniere saisie"
        '
        'last_numero
        '
        Me.last_numero.AutoSize = True
        Me.last_numero.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.last_numero.Location = New System.Drawing.Point(183, 377)
        Me.last_numero.Name = "last_numero"
        Me.last_numero.Size = New System.Drawing.Size(79, 13)
        Me.last_numero.TabIndex = 6
        Me.last_numero.Text = "LLLL99999999"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(427, 432)
        Me.Controls.Add(Me.last_numero)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.resultat)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label_resultat)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label_resultat As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents resultat As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents last_numero As System.Windows.Forms.Label

End Class

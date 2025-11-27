Imports System.Drawing.Printing
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Imports System.Threading.Tasks
Imports System.Globalization

Public Class FrmProformaInvoice
    Inherits Form

    ' === Controls ===
    Private PanelHeader As Panel
    Private contentPanel As Panel
    Private lblCompanyName, lblCompanyDetails, lblInvoiceTitle As Label
    Private cmbInvoiceType As ComboBox
    Private dgvInvoiceItems As DataGridView
    Private txtTotalCost As TextBox
    Private lblTotalCost As Label
    Private btnPrint As Button
    Private lblNote, lblThanks As Label
    Private txtNote As TextBox
    Private txtThanks As TextBox
    Private lblBilledTo, lblAddress, lblInvoiceDate, lblInvoiceSerial As Label
    Private txtBilledTo, txtAddress, txtInvoiceSerial As TextBox
    Private dtpInvoiceDate As DateTimePicker
    Private PrintDocument1 As PrintDocument
    Private PrintPreviewDialog1 As PrintPreviewDialog

    ' New product entry controls
    Private grpProductEntry As GroupBox
    Private lblItemNo As Label
    Private txtItemNo As TextBox
    Private lblDescription As Label
    Private txtDescription As TextBox
    Private lblQty As Label
    Private txtQty As TextBox
    Private lblUnitPrice As Label
    Private txtUnitPrice As TextBox
    Private btnAddItem As Button

    ' New footer controls
    Private btnRemoveLine As Button
    Private btnResetAll As Button
    Private btnLoadLicense As Button
    Private btnLicenseStatus As Button
    Private btnSaveInvoice As Button
    Private btnViewHistory As Button
    Private pnlEditInvoice As Panel
    Private btnSaveChanges As Button
    Private editingInvoiceId As Integer = -1

    ' Client ID UI
    Private lblClientIdLabel As Label
    Private lblClientIdValue As Label

    ' Trial status UI
    Private lblTrialStatus As Label

    ' UI helpers
    Private toolTip1 As ToolTip
    Private btnPrintDefaultBackColor As Color

    ' Local client identifier
    Private clientId As String

    ' Timer to update trial label daily
    Private trialTimer As Timer

    ' Serial counter file name (persisted in AppData)
    Private Const SerialCounterFileName As String = "serial_counter.txt"

    Public Sub New()
        InitializeComponent()
        EnsureClientId()
        UpdateClientIdDisplay()
        UpdateTrialLabel()
        UpdatePrintButtonState()
        AddHandler Me.Load, AddressOf FrmProformaInvoice_Load
    End Sub

    Private Sub InitializeComponent()
        ' === Form ===
        Me.Text = "Invoice Generator"
        Me.ClientSize = New Size(1000, 750)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.BackColor = Color.White

        ' === Header Panel (docked) ===
        PanelHeader = New Panel With {
            .Dock = DockStyle.Top,
            .Height = 120,
            .BackColor = Color.White
        }

        lblCompanyName = New Label With {
            .Text = "MWASID GRANITES",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .Location = New Point(20, 10),
            .AutoSize = True
        }

        lblCompanyDetails = New Label With {
            .Text = "Nairobi, 00208, Kenya" & vbCrLf &
                    "Contacts: 0790109743  |  Email: mwasidgranite@gmail.com",
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(20, 45),
            .AutoSize = True
        }

        lblInvoiceTitle = New Label With {
            .Text = "PROFORMA INVOICE",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold Or FontStyle.Underline),
            .Location = New Point(720, 40),
            .AutoSize = True
        }

        ' small labels to show the ClientID
        lblClientIdLabel = New Label With {
            .Text = "Client ID:",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .Location = New Point(520, 70),
            .AutoSize = True
        }
        lblClientIdValue = New Label With {
            .Text = "Generating...",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .Location = New Point(580, 70),
            .AutoSize = True
        }

        ' trial status label (below client id)
        lblTrialStatus = New Label With {
            .Text = "",
            .Font = New Font("Segoe UI", 8, FontStyle.Italic),
            .Location = New Point(520, 90),
            .AutoSize = True,
            .ForeColor = Color.DarkRed
        }

        ' ComboBox to allow editing/selecting invoice type; placed in header
        cmbInvoiceType = New ComboBox With {
            .Location = New Point(520, 40),
            .Width = 180,
            .DropDownStyle = ComboBoxStyle.DropDown ' allow typing custom value
        }
        cmbInvoiceType.Items.AddRange(New Object() {"PROFORMA INVOICE", "TAX INVOICE", "QUOTATION", "CREDIT NOTE", "DEBIT NOTE", "CASH SALE"})
        cmbInvoiceType.Text = lblInvoiceTitle.Text
        AddHandler cmbInvoiceType.TextChanged, AddressOf cmbInvoiceType_TextChanged

        ' Save Invoice button in header (top)
        btnSaveInvoice = New Button With {
            .Text = "Save Invoice",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(820, 680), ' moved to footer area
            .Size = New Size(120, 28),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSaveInvoice.Click, AddressOf btnSaveInvoice_Click

        PanelHeader.Controls.AddRange({lblCompanyName, lblCompanyDetails, lblInvoiceTitle, lblClientIdLabel, lblClientIdValue, lblTrialStatus})
        PanelHeader.Controls.Add(cmbInvoiceType)

        ' === Content panel with AutoScroll to allow overlapping elements ===
        contentPanel = New Panel With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .BackColor = Color.Transparent
        }

        ' === Product Entry Group (will be placed just above the grid) ===
        grpProductEntry = New GroupBox With {
            .Location = New Point(20, 200), ' moved to sit immediately above the DataGridView
            .Size = New Size(960, 50),
            .FlatStyle = FlatStyle.Flat
        }

        lblItemNo = New Label With {.Text = "Item #", .Location = New Point(10, 20), .AutoSize = True}
        txtItemNo = New TextBox With {.Location = New Point(60, 16), .Width = 60}

        lblDescription = New Label With {.Text = "Description", .Location = New Point(130, 20), .AutoSize = True}
        txtDescription = New TextBox With {.Location = New Point(210, 16), .Width = 360}

        lblQty = New Label With {.Text = "Qty", .Location = New Point(580, 20), .AutoSize = True}
        txtQty = New TextBox With {.Location = New Point(610, 16), .Width = 60}

        lblUnitPrice = New Label With {.Text = "Unit Price", .Location = New Point(690, 20), .AutoSize = True}
        txtUnitPrice = New TextBox With {.Location = New Point(760, 16), .Width = 90}

        btnAddItem = New Button With {
            .Text = "Add Item",
            .Location = New Point(860, 13),
            .Size = New Size(80, 25),
            .BackColor = Color.FromArgb(52, 73, 94),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAddItem.Click, AddressOf btnAddItem_Click

        grpProductEntry.Controls.AddRange({lblItemNo, txtItemNo, lblDescription, txtDescription, lblQty, txtQty, lblUnitPrice, txtUnitPrice, btnAddItem})

        ' === Client Info (moved to top, above the product entry group) ===
        lblBilledTo = New Label With {
            .Text = "Bill To:",
            .Location = New Point(20, 130),
            .Font = New Font("Segoe UI", 10),
            .AutoSize = True,
            .TextAlign = ContentAlignment.MiddleRight
        }
        txtBilledTo = New TextBox With {
            .Location = New Point(100, 128),
            .Size = New Size(360, 24),
            .Multiline = False,
            .BorderStyle = BorderStyle.FixedSingle,
            .Font = New Font("Segoe UI", 9),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
        }

        lblInvoiceDate = New Label With {.Text = "Invoice Date:", .Location = New Point(480, 130), .Font = New Font("Segoe UI", 10)}
        dtpInvoiceDate = New DateTimePicker With {.Location = New Point(600, 128), .Width = 150}

        lblAddress = New Label With {
            .Text = "Address:",
            .Location = New Point(20, 160),
            .Font = New Font("Segoe UI", 10),
            .AutoSize = True,
            .TextAlign = ContentAlignment.MiddleRight
        }
        txtAddress = New TextBox With {
            .Location = New Point(100, 158),
            .Size = New Size(360, 24),
            .Multiline = False,
            .BorderStyle = BorderStyle.FixedSingle,
            .Font = New Font("Segoe UI", 9),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
        }

        lblInvoiceSerial = New Label With {.Text = "Invoice Serial:", .Location = New Point(480, 160), .Font = New Font("Segoe UI", 10)}
        txtInvoiceSerial = New TextBox With {.Location = New Point(600, 158), .Width = 150}

        ' === DataGridView (Invoice Table) ===
        dgvInvoiceItems = New DataGridView With {
            .Location = New Point(20, 260), ' placed below the product entry group
            .Size = New Size(960, 260),
            .BackgroundColor = Color.White,
            .AllowUserToAddRows = False,
            .RowHeadersVisible = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }

        dgvInvoiceItems.Columns.Add("ItemNo", "ITEM NO.")
        dgvInvoiceItems.Columns.Add("Description", "DESCRIPTION")
        dgvInvoiceItems.Columns.Add("Qty", "QTY")
        dgvInvoiceItems.Columns.Add("UnitPrice", "UNIT PRICE (KSH)")
        dgvInvoiceItems.Columns.Add("Amount", "T. AMOUNT (KSH)")

        ' === Total Section ===
        lblTotalCost = New Label With {
            .Text = "Total Cost (KES):",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Location = New Point(650, 520),
            .AutoSize = True
        }

        txtTotalCost = New TextBox With {
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Location = New Point(790, 516),
            .Width = 180,
            .TextAlign = HorizontalAlignment.Right,
            .Text = "0.00",
            .ReadOnly = True
        }

        ' === Note and Thank You (editable) ===
        lblNote = New Label With {
            .Text = "Note:",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(20, 560),
            .AutoSize = True
        }
        txtNote = New TextBox With {
            .Location = New Point(70, 556),
            .Size = New Size(700, 40),
            .Multiline = True,
            .Font = New Font("Segoe UI", 9),
            .Text = "All the Logistics and Transport Cost are included in the total cost."
        }

        lblThanks = New Label With {
            .Text = "Thank you message:",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(20, 606),
            .AutoSize = True
        }
        txtThanks = New TextBox With {
            .Location = New Point(140, 602),
            .Size = New Size(630, 40),
            .Multiline = True,
            .Font = New Font("Segoe UI", 9),
            .Text = "Thank you for your valuable inquiry."
        }

        ' === Footer Buttons: Remove Line, Reset All, Print ===
        btnRemoveLine = New Button With {
            .Text = "Remove Line",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(20, 680),
            .Size = New Size(100, 28),
            .BackColor = Color.FromArgb(52, 73, 94),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnRemoveLine.Click, AddressOf btnRemoveLine_Click

        btnResetAll = New Button With {
            .Text = "Reset All",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(140, 680),
            .Size = New Size(100, 28),
            .BackColor = Color.FromArgb(52, 73, 94),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnResetAll.Click, AddressOf btnResetAll_Click

        btnViewHistory = New Button With {
            .Text = "View History",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(260, 680),
            .Size = New Size(120, 28),
            .BackColor = Color.FromArgb(41, 128, 185),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnViewHistory.Click, AddressOf btnViewHistory_Click

        btnLicenseStatus = New Button With {
            .Text = "License Status",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(400, 680),
            .Size = New Size(120, 28),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnLicenseStatus.Click, AddressOf btnLicenseStatus_Click

        btnLoadLicense = New Button With {
            .Text = "Load License",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(540, 680),
            .Size = New Size(120, 28),
            .BackColor = Color.FromArgb(46, 204, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnLoadLicense.Click, AddressOf btnLoadLicense_Click

        btnSaveInvoice = New Button With {
            .Text = "Save Invoice",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .Location = New Point(680, 680),
            .Size = New Size(120, 28),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSaveInvoice.Click, AddressOf btnSaveInvoice_Click

        btnPrint = New Button With {
            .Text = "Print",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Location = New Point(880, 680),
            .Size = New Size(80, 30),
            .BackColor = Color.FromArgb(52, 73, 94),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Enabled = False ' disabled until there is at least one item and UI is enabled
        }
        AddHandler btnPrint.Click, AddressOf btnPrint_Click

        ' store default print button color and tooltip
        btnPrintDefaultBackColor = btnPrint.BackColor
        toolTip1 = New ToolTip()
        toolTip1.SetToolTip(btnPrint, "Add at least one item to enable printing")

        ' === Print Setup ===
        PrintDocument1 = New PrintDocument()
        PrintPreviewDialog1 = New PrintPreviewDialog() With {.Document = PrintDocument1, .Width = 800, .Height = 600}
        AddHandler PrintDocument1.PrintPage, AddressOf PrintDocument1_PrintPage

        ' === Add Controls to contentPanel ===
        contentPanel.Controls.AddRange({lblBilledTo, txtBilledTo, lblAddress, txtAddress, lblInvoiceDate, dtpInvoiceDate, lblInvoiceSerial, txtInvoiceSerial, grpProductEntry, dgvInvoiceItems, lblTotalCost, txtTotalCost, lblNote, txtNote, lblThanks, txtThanks, btnRemoveLine, btnResetAll, btnViewHistory, btnLicenseStatus, btnLoadLicense, btnSaveInvoice, btnPrint})

        ' === Add header and contentPanel to Form ===
        Me.Controls.AddRange({PanelHeader, contentPanel})

        ' === Calculation Events ===
        AddHandler dgvInvoiceItems.CellValueChanged, AddressOf dgvInvoiceItems_CellValueChanged
        AddHandler dgvInvoiceItems.RowsAdded, AddressOf dgvInvoiceItems_RowsAdded
        AddHandler dgvInvoiceItems.UserDeletedRow, AddressOf dgvInvoiceItems_UserDeletedRow
        AddHandler dgvInvoiceItems.RowsRemoved, AddressOf dgvInvoiceItems_RowsRemoved

        ' === Trial update timer (checks once per day) ===
        trialTimer = New Timer()
        ' 24 hours in milliseconds
        trialTimer.Interval = 24 * 60 * 60 * 1000
        AddHandler trialTimer.Tick, AddressOf TrialTimer_Tick
        trialTimer.Start()
        ' Ensure initial print button state
        UpdatePrintButtonState()

        ' Generate initial invoice serial (persisted auto-increment + GUID)
        SetNewInvoiceSerial()
    End Sub

    ' Generate and set a new invoice serial using a persisted auto-increment counter plus a GUID.
    Private Sub SetNewInvoiceSerial()
        Try
            LicenseManager.EnsureAppFolder()
            Dim counterFile = Path.Combine(LicenseManager.AppFolder, SerialCounterFileName)
            Dim counter As Long = 0
            If File.Exists(counterFile) Then
                Long.TryParse(File.ReadAllText(counterFile).Trim(), counter)
            End If
            counter += 1
            File.WriteAllText(counterFile, counter.ToString())

            Dim guidPart = Guid.NewGuid().ToString("N").ToUpper()
            ' Format: INV-000123-<GUID>
            Dim serial = String.Format("INV-{0}-{1}", counter.ToString("D6"), guidPart)
            If txtInvoiceSerial IsNot Nothing Then txtInvoiceSerial.Text = serial
        Catch
            ' ignore errors - do not crash UI
        End Try
    End Sub

    ' === Auto Calculation ===
    Private Sub RecalculateTotal()
        Dim total As Decimal = 0
        For Each row As DataGridViewRow In dgvInvoiceItems.Rows
            If Not row.IsNewRow Then
                Dim qty, price As Decimal
                Decimal.TryParse(Convert.ToString(row.Cells("Qty").Value), qty)
                Decimal.TryParse(Convert.ToString(row.Cells("UnitPrice").Value), price)
                Dim amount As Decimal = qty * price
                row.Cells("Amount").Value = amount.ToString("N2")
                total += amount
            End If
        Next
        txtTotalCost.Text = total.ToString("N2")
        UpdatePrintButtonState()
    End Sub

    Private Sub dgvInvoiceItems_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then RecalculateTotal()
    End Sub

    Private Sub dgvInvoiceItems_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)
        RecalculateTotal()
    End Sub

    Private Sub dgvInvoiceItems_UserDeletedRow(sender As Object, e As DataGridViewRowEventArgs)
        RecalculateTotal()
    End Sub

    Private Sub dgvInvoiceItems_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs)
        RecalculateTotal()
    End Sub

    ' === Add Item button handler ===
    Private Sub btnAddItem_Click(sender As Object, e As EventArgs)
        ' Try to parse values
        Dim itemNoText = txtItemNo.Text.Trim()
        Dim description = txtDescription.Text.Trim()
        Dim qty As Decimal = 0D
        Dim unitPrice As Decimal = 0D
        Decimal.TryParse(txtQty.Text.Trim(), qty)
        Decimal.TryParse(txtUnitPrice.Text.Trim(), unitPrice)

        If String.IsNullOrWhiteSpace(description) Then
            MessageBox.Show("Please enter a description.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If qty <= 0 OrElse unitPrice < 0 Then
            MessageBox.Show("Please enter valid quantity and unit price.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim amount = qty * unitPrice

        Dim itemNo As String = itemNoText
        If String.IsNullOrWhiteSpace(itemNo) Then
            itemNo = (dgvInvoiceItems.Rows.Count + 1).ToString()
        End If
        dgvInvoiceItems.Rows.Add(itemNo, description, qty.ToString(), unitPrice.ToString("N2"), amount.ToString("N2"))
        RecalculateTotal()
        ' Clear inputs
        txtItemNo.Clear()
        txtDescription.Clear()
        txtQty.Clear()
        txtUnitPrice.Clear()
        txtDescription.Focus()
    End Sub

    ' === Remove Line ===
    Private Sub btnRemoveLine_Click(sender As Object, e As EventArgs)
        If dgvInvoiceItems.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In dgvInvoiceItems.SelectedRows
                If Not r.IsNewRow Then dgvInvoiceItems.Rows.Remove(r)
            Next
            RecalculateTotal()
        Else
            MessageBox.Show("Select a row to remove.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' === Reset All ===
    Private Sub btnResetAll_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("Clear all invoice items?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            dgvInvoiceItems.Rows.Clear()
            txtTotalCost.Text = "0.00"
            UpdatePrintButtonState()
            ' Reset invoice serial when starting a fresh invoice
            SetNewInvoiceSerial()
        End If
    End Sub

    ' === Print Logic ===
    Private Sub btnPrint_Click(sender As Object, e As EventArgs)
        ' Prevent printing when there are no items
        Dim hasItems As Boolean = False
        For Each r As DataGridViewRow In dgvInvoiceItems.Rows
            If Not r.IsNewRow Then
                hasItems = True
                Exit For
            End If
        Next
        If Not hasItems Then
            MessageBox.Show("No items to print.", "Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
            UpdatePrintButtonState()
            Return
        End If

        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs)
        Dim g As Graphics = e.Graphics
        Dim fontHeader As New Font("Segoe UI", 12, FontStyle.Bold)
        Dim fontNormal As New Font("Segoe UI", 10)
        Dim y As Integer = 50

        ' Header
        g.DrawString(lblCompanyName.Text, fontHeader, Brushes.Black, 50, y) : y += 25
        g.DrawString(lblCompanyDetails.Text, fontNormal, Brushes.Black, 50, y) : y += 40
        ' print the dynamic invoice title from lblInvoiceTitle
        g.DrawString(lblInvoiceTitle.Text, fontHeader, Brushes.Black, 600, 50)
        y += 10

        ' Client Info (moved to top)
        g.DrawString("Billed To: " & txtBilledTo.Text, fontNormal, Brushes.Black, 50, y)
        g.DrawString("Invoice Date: " & dtpInvoiceDate.Value.ToShortDateString(), fontNormal, Brushes.Black, 600, y)
        y += 20
        g.DrawString("Address: " & txtAddress.Text, fontNormal, Brushes.Black, 50, y)
        g.DrawString("Invoice Serial: " & txtInvoiceSerial.Text, fontNormal, Brushes.Black, 600, y)
        y += 30

        ' Table Header
        g.DrawString("ITEM NO.", fontNormal, Brushes.Black, 50, y)
        g.DrawString("DESCRIPTION", fontNormal, Brushes.Black, 120, y)
        g.DrawString("QTY", fontNormal, Brushes.Black, 400, y)
        g.DrawString("UNIT PRICE", fontNormal, Brushes.Black, 480, y)
        g.DrawString("AMOUNT", fontNormal, Brushes.Black, 600, y)
        y += 20

        ' Items
        For Each row As DataGridViewRow In dgvInvoiceItems.Rows
            If Not row.IsNewRow Then
                g.DrawString(Convert.ToString(row.Cells("ItemNo").Value), fontNormal, Brushes.Black, 50, y)
                g.DrawString(Convert.ToString(row.Cells("Description").Value), fontNormal, Brushes.Black, 120, y)
                g.DrawString(Convert.ToString(row.Cells("Qty").Value), fontNormal, Brushes.Black, 400, y)
                g.DrawString(Convert.ToString(row.Cells("UnitPrice").Value), fontNormal, Brushes.Black, 480, y)
                g.DrawString(Convert.ToString(row.Cells("Amount").Value), fontNormal, Brushes.Black, 600, y)
                y += 20
            End If
        Next

        y += 30
        g.DrawString("Total Cost (KES): " & txtTotalCost.Text, fontHeader, Brushes.Black, 480, y)
        y += 40
        ' Print editable note and thank-you message (use textboxes' content)
        g.DrawString(Convert.ToString(txtNote.Text), fontNormal, Brushes.Black, 50, y) : y += 20
        g.DrawString(Convert.ToString(txtThanks.Text), fontNormal, Brushes.Black, 50, y)
    End Sub

    ' Sync ComboBox text into the header invoice title label
    Private Sub cmbInvoiceType_TextChanged(sender As Object, e As EventArgs)
        If cmbInvoiceType IsNot Nothing AndAlso lblInvoiceTitle IsNot Nothing Then
            lblInvoiceTitle.Text = cmbInvoiceType.Text
        End If
    End Sub

    ' Ensure a local client identifier exists and is stored in AppData
    Private Sub EnsureClientId()
        Try
            Dim appFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "InvoiceGenerator")
            If Not Directory.Exists(appFolder) Then Directory.CreateDirectory(appFolder)
            Dim clientFile = Path.Combine(appFolder, "clientid.txt")
            If File.Exists(clientFile) Then
                Dim content = File.ReadAllText(clientFile).Trim()
                If Not String.IsNullOrEmpty(content) Then
                    clientId = content
                    Return
                End If
            End If

            ' create and persist a new GUID
            clientId = Guid.NewGuid().ToString()
            File.WriteAllText(clientFile, clientId)
        Catch ex As Exception
            ' If writing fails, keep clientId empty - do not crash the form
            clientId = String.Empty
        End Try
    End Sub

    ' Update the client id label in UI
    Private Sub UpdateClientIdDisplay()
        Try
            If lblClientIdValue IsNot Nothing Then
                If String.IsNullOrEmpty(clientId) Then
                    lblClientIdValue.Text = "N/A"
                Else
                    lblClientIdValue.Text = clientId
                End If
            End If
        Catch
            ' ignore
        End Try
    End Sub

    ' Update the trial/license status label
    Private Sub UpdateTrialLabel()
        Try
            If lblTrialStatus Is Nothing Then Return
            LicenseManager.EnsureAppFolder()

            If LicenseManager.IsTrialActive() Then
                Dim daysLeft = LicenseManager.TrialDaysLeft()
                lblTrialStatus.Text = "Free trial — " & daysLeft & " day(s) left"
                lblTrialStatus.ForeColor = Color.DarkGreen
                Return
            End If

            ' If not trial, check license file and show expiry or status
            Dim expiry As DateTime = DateTime.MinValue
            Dim licensedClient As String = String.Empty
            If LicenseManager.TryValidateLicense(expiry, licensedClient) Then
                Dim display = expiry.ToLocalTime().ToString("yyyy-MM-dd")
                If Not String.IsNullOrEmpty(clientId) AndAlso String.Equals(clientId, licensedClient, StringComparison.OrdinalIgnoreCase) Then
                    lblTrialStatus.Text = "Licensed until " & display
                    lblTrialStatus.ForeColor = Color.DarkBlue
                Else
                    lblTrialStatus.Text = "License installed (for other client) until " & display
                    lblTrialStatus.ForeColor = Color.OrangeRed
                End If
            Else
                lblTrialStatus.Text = "Trial expired — license required"
                lblTrialStatus.ForeColor = Color.DarkRed
            End If
        Catch
            ' ignore UI update errors
        End Try
    End Sub

    ' Timer tick: update trial label and enforce inactive state when needed
    Private Sub TrialTimer_Tick(sender As Object, e As EventArgs)
        Try
            ' Refresh label
            UpdateTrialLabel()

            ' If trial expired and not licensed, disable UI
            If Not LicenseManager.IsTrialActive() Then
                Dim expiryUtc As DateTime = DateTime.MinValue
                If Not LicenseManager.IsLicensed(clientId, expiryUtc) Then
                    ' ensure UI is disabled and inform user
                    ApplyLicensedState(False)
                    ' notify once (avoid spamming) - show a simple message box
                    MessageBox.Show("Trial expired. Please install a valid license to continue using the application.", "License Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        Catch
            ' ignore timer errors
        End Try
    End Sub

    ' === License / Trial handling on load ===
    Private Sub FrmProformaInvoice_Load(sender As Object, e As EventArgs)
        Try
            LicenseManager.EnsureAppFolder()
            If LicenseManager.IsTrialActive() Then
                Dim daysLeft = LicenseManager.TrialDaysLeft()
                If daysLeft <= 7 Then
                    MessageBox.Show("Trial will expire in " & daysLeft &
                                    " day(s). After that a signed license file will be required.", "Trial Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                ApplyLicensedState(True)
                Return
            End If

            ' Trial expired - check license matching client id
            Dim expiryUtc As DateTime = DateTime.MinValue
            If LicenseManager.IsLicensed(clientId, expiryUtc) Then
                ApplyLicensedState(True)
                ' License is valid: switch to Pro title and hide license UI/timer
                Me.Text = "invoice generator Pro"
                ' ensure trial is ended if license already present
                Try
                    LicenseManager.EndTrial()
                Catch
                End Try
                If btnLoadLicense IsNot Nothing Then btnLoadLicense.Visible = False
                If lblTrialStatus IsNot Nothing Then lblTrialStatus.Visible = False
                If trialTimer IsNot Nothing Then trialTimer.Enabled = False
                Return
            End If

            ' not licensed
            MessageBox.Show("Trial expired. Place a signed license file named 'license.lic' in: " & LicenseManager.AppFolder & vbCrLf & vbCrLf & "Contact the vendor to obtain a signed license (RSA-signed payload containing an expiry date and matching ClientID).", "License Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ApplyLicensedState(False)
        Catch ex As Exception
            MessageBox.Show("License check failed: " & ex.Message, "License Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ApplyLicensedState(False)
        End Try
    End Sub

    Private Sub ApplyLicensedState(enabled As Boolean)
        contentPanel.Enabled = enabled
        PanelHeader.Enabled = enabled
        btnAddItem.Enabled = enabled
        btnRemoveLine.Enabled = enabled
        btnResetAll.Enabled = enabled
        ' keep print enabled only if UI licensed and there are items
        If enabled Then
            UpdatePrintButtonState()
        Else
            btnPrint.Enabled = False
        End If
        dgvInvoiceItems.Enabled = enabled
    End Sub

    ' Load or paste a license file, save to AppData and validate
    Private Sub btnLoadLicense_Click(sender As Object, e As EventArgs)
        Try
            LicenseManager.EnsureAppFolder()

            ' First allow pasting via InputBox
            Dim pastePrompt = "Paste the license content here (two lines: payload and base64 signature). Leave empty to pick a file..."
            Dim pasted = Interaction.InputBox(pastePrompt, "Paste License", "")
            Dim licenseText As String = String.Empty

            If Not String.IsNullOrWhiteSpace(pasted) Then
                licenseText = pasted
            Else
                Using ofd As New OpenFileDialog()
                    ofd.Filter = "License files (*.lic)|*.lic|All files (*.*)|*.*"
                    ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    If ofd.ShowDialog() = DialogResult.OK Then
                        licenseText = File.ReadAllText(ofd.FileName)
                    Else
                        Return
                    End If
                End Using
            End If

            If String.IsNullOrWhiteSpace(licenseText) Then
                MessageBox.Show("No license content provided.", "Load License", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim target = Path.Combine(LicenseManager.AppFolder, "license.lic")
            File.WriteAllText(target, licenseText)

            Dim expiryUtc As DateTime = DateTime.MinValue
            If LicenseManager.IsLicensed(clientId, expiryUtc) Then
                MessageBox.Show("License installed. Expires: " & expiryUtc.ToLocalTime().ToString("yyyy-MM-dd"), "License Installed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ApplyLicensedState(True)
                UpdateTrialLabel()
                UpdatePrintButtonState()
                ' On successful license, update UI to Pro and hide license controls + disable timer
                Me.Text = "invoice generator Pro"
                If btnLoadLicense IsNot Nothing Then btnLoadLicense.Visible = False
                If lblTrialStatus IsNot Nothing Then lblTrialStatus.Visible = False
                If trialTimer IsNot Nothing Then trialTimer.Enabled = False
                ' End trial now that license is installed
                Try
                    LicenseManager.EndTrial()
                Catch
                End Try
            Else
                MessageBox.Show("License is invalid or does not match this Client ID.", "Invalid License", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ApplyLicensedState(False)
            End If
        Catch ex As Exception
            MessageBox.Show("Failed to install license: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Show license/trial status dialog
    Private Sub btnLicenseStatus_Click(sender As Object, e As EventArgs)
        Try
            LicenseManager.EnsureAppFolder()
            Dim expiry As DateTime = DateTime.MinValue
            Dim licensedClient As String = String.Empty
            Dim msg As String = String.Empty

            If LicenseManager.TryValidateLicense(expiry, licensedClient) Then
                Dim localMatch = If(String.IsNullOrEmpty(clientId), False, String.Equals(clientId, licensedClient, StringComparison.OrdinalIgnoreCase))
                Dim daysLeft As Integer = CInt((expiry.ToUniversalTime() - DateTime.UtcNow).TotalDays)
                If daysLeft < 0 Then daysLeft = 0
                msg &= "Licensed: Yes" & vbCrLf
                msg &= "Expires (UTC): " & expiry.ToUniversalTime().ToString("yyyy-MM-dd") & vbCrLf
                msg &= "Days until expiry: " & daysLeft & vbCrLf
                msg &= "Licensed Client ID: " & licensedClient & vbCrLf
                msg &= "This device Client ID: " & If(String.IsNullOrEmpty(clientId), "N/A", clientId) & vbCrLf
                msg &= "Client ID matches license: " & If(localMatch, "Yes", "No") & vbCrLf
            Else
                ' No valid license file present
                If File.Exists(Path.Combine(LicenseManager.AppFolder, "license.lic")) Then
                    msg &= "License file present but invalid or signature failed." & vbCrLf
                Else
                    msg &= "No license installed." & vbCrLf
                End If
                If LicenseManager.IsTrialActive() Then
                    msg &= "Trial active. Days left: " & LicenseManager.TrialDaysLeft() & vbCrLf
                Else
                    msg &= "Trial expired." & vbCrLf
                End If
            End If

            MessageBox.Show(msg, "License Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Failed to determine license status: " & ex.Message, "License Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Enable/disable Print button depending on whether there are items in the grid and UI state
    Private Sub UpdatePrintButtonState()
        Try
            If btnPrint Is Nothing OrElse dgvInvoiceItems Is Nothing Then
                Return
            End If

            Dim hasItems As Boolean = False
            For Each r As DataGridViewRow In dgvInvoiceItems.Rows
                If Not r.IsNewRow Then
                    hasItems = True
                    Exit For
                End If
            Next

            ' When the overall UI is disabled (trial expired / not licensed) keep the button disabled
            Dim uiEnabled = True
            Try
                uiEnabled = If(contentPanel IsNot Nothing, contentPanel.Enabled, True)
            Catch
            End Try

            If Not uiEnabled Then
                btnPrint.Enabled = False
                btnPrint.BackColor = Color.Gray
                btnPrint.ForeColor = Color.LightGray
                btnPrint.Cursor = Cursors.Default
                If toolTip1 IsNot Nothing Then toolTip1.SetToolTip(btnPrint, "Application not licensed or trial expired")
                Return
            End If

            ' UI is enabled: set appearance based on whether there are items
            btnPrint.Enabled = True
            If hasItems Then
                btnPrint.BackColor = btnPrintDefaultBackColor
                btnPrint.ForeColor = Color.White
                btnPrint.Cursor = Cursors.Default
                If toolTip1 IsNot Nothing Then toolTip1.SetToolTip(btnPrint, "Print preview and print the invoice")
            Else
                btnPrint.BackColor = Color.Gray
                btnPrint.ForeColor = Color.LightGray
                btnPrint.Cursor = Cursors.No
                If toolTip1 IsNot Nothing Then toolTip1.SetToolTip(btnPrint, "Add at least one item to enable printing")
            End If

            ' Also disable/enable PrintPreviewDialog controls (ToolStrip and PreviewControl)
            Try
                If PrintPreviewDialog1 IsNot Nothing Then
                    If PrintPreviewDialog1.PrintPreviewControl IsNot Nothing Then
                        PrintPreviewDialog1.PrintPreviewControl.Enabled = hasItems
                    End If
                    For Each c As Control In PrintPreviewDialog1.Controls
                        If TypeOf c Is ToolStrip Then
                            c.Enabled = hasItems
                        End If
                    Next
                End If
            Catch
                ' ignore
            End Try
        Catch
            ' ignore
        End Try
    End Sub

    ' --- Database helper functions ---
    Private Function GetConnectionString() As String
        Dim dbPath = Path.Combine(LicenseManager.AppFolder, "invoices.accdb")
        Return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;"
    End Function

    ' Ensure the Access database file and required tables exist. If the .accdb file is missing
    ' attempt to create it via ADOX (late-bound). Then ensure the Invoices and InvoiceItems
    ' tables exist and create them if not present.
    Private Sub EnsureDatabase()
        Try
            LicenseManager.EnsureAppFolder()
            Dim dbPath = Path.Combine(LicenseManager.AppFolder, "invoices.accdb")

            ' If the DB file does not exist, attempt to create it using ADOX (late-bound COM).
            ' This requires the Access Database Engine (ACE) to be installed on the system.
            If Not File.Exists(dbPath) Then
                Try
                    Dim cat = CreateObject("ADOX.Catalog")
                    ' Create with ACE provider; use concatenation to avoid VB string-interpolation issues.
                    cat.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";")
                    cat = Nothing
                Catch ex As Exception
                    ' Creation may fail if ADOX/ACE is not available. We'll surface a helpful
                    ' debug trace but not throw to avoid crashing the UI thread here.
                    Try
                        ' Best effort: write a small log file to help diagnose.
                        Dim logPath = Path.Combine(LicenseManager.AppFolder, "db_create_error.log")
                        File.AppendAllText(logPath, DateTime.UtcNow.ToString("s") & " - ADOX create failed: " & ex.Message & vbCrLf)
                    Catch
                    End Try
                End Try
            End If

            ' If the file exists (either previously or created above), ensure the tables exist.
            If File.Exists(dbPath) Then
                Dim connStr = GetConnectionString()
                Using conn As New OleDbConnection(connStr)
                    conn.Open()

                    ' Query schema for tables
                    Dim schema = conn.GetSchema("Tables")
                    Dim hasInvoices As Boolean = False
                    Dim hasItems As Boolean = False
                    For Each r As DataRow In schema.Rows
                        Dim tn = Convert.ToString(r("TABLE_NAME"))
                        Dim tt = Convert.ToString(r("TABLE_TYPE"))
                        If String.Equals(tt, "TABLE", StringComparison.OrdinalIgnoreCase) Then
                            If String.Equals(tn, "Invoices", StringComparison.OrdinalIgnoreCase) Then hasInvoices = True
                            If String.Equals(tn, "InvoiceItems", StringComparison.OrdinalIgnoreCase) Then hasItems = True
                        End If
                    Next

                    ' Create Invoices table if missing
                    If Not hasInvoices Then
                        ' AUTOINCREMENT (COUNTER) primary key; TEXT(255) for strings, DATETIME for dates, DOUBLE for totals
                        Dim createInvoices As String = "CREATE TABLE Invoices (ID COUNTER PRIMARY KEY, InvoiceSerial TEXT(255), InvoiceDate DATETIME, Client TEXT(255), Total DOUBLE)"
                        Using cmd As New OleDbCommand(createInvoices, conn)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If

                    ' Create InvoiceItems table if missing
                    If Not hasItems Then
                        Dim createItems As String = "CREATE TABLE InvoiceItems (ID COUNTER PRIMARY KEY, InvoiceID LONG, ItemDescription TEXT(255), Qty DOUBLE, UnitPrice DOUBLE, Amount DOUBLE)"
                        Using cmd As New OleDbCommand(createItems, conn)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                End Using
            End If
        Catch ex As Exception
            ' Log the error silently to help troubleshooting but do not crash the application.
            Try
                Dim logPath = Path.Combine(LicenseManager.AppFolder, "db_error.log")
                File.AppendAllText(logPath, DateTime.UtcNow.ToString("s") & " - EnsureDatabase error: " & ex.Message & vbCrLf)
            Catch
            End Try
        End Try
    End Sub

    ' Save invoice (header + items) to Access DB. Uses background task.
    Private Sub btnSaveInvoice_Click(sender As Object, e As EventArgs)
        If Not LicenseManager.IsLicenseValid(clientId) Then
            MessageBox.Show("Saving invoices requires a valid license.", "License Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        SaveInvoice()
    End Sub

    Private Sub SaveInvoice()
        Task.Run(Sub()
                     Try
                         EnsureDatabase()
                         Dim connStr = GetConnectionString()
                         Using conn As New OleDbConnection(connStr)
                             conn.Open()
                             ' Validate header fields on UI thread
                             Dim billedTo = String.Empty
                             Dim invoiceDate As DateTime = DateTime.UtcNow
                             Dim total As Decimal = 0D
                             Me.Invoke(Sub()
                                           billedTo = txtBilledTo.Text.Trim()
                                           invoiceDate = dtpInvoiceDate.Value
                                           Decimal.TryParse(txtTotalCost.Text, NumberStyles.Number, CultureInfo.CurrentCulture, total)
                                       End Sub)
                             If String.IsNullOrWhiteSpace(billedTo) Then
                                 Me.Invoke(Sub() MessageBox.Show("Client name is required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning))
                                 Return
                             End If
                             ' Insert header - explicit parameter types and strongly-typed values
                             Dim cmd As New OleDbCommand("INSERT INTO Invoices (InvoiceSerial, InvoiceDate, Client, Total) VALUES (?, ?, ?, ?)", conn)
                             cmd.Parameters.Add("p1", OleDbType.VarWChar).Value = If(txtInvoiceSerial IsNot Nothing, txtInvoiceSerial.Text, "")
                             cmd.Parameters.Add("p2", OleDbType.Date).Value = invoiceDate
                             cmd.Parameters.Add("p3", OleDbType.VarWChar).Value = billedTo
                             cmd.Parameters.Add("p4", OleDbType.Double).Value = Convert.ToDouble(total)
                             cmd.ExecuteNonQuery()
                             ' Get generated ID
                             Dim idCmd As New OleDbCommand("SELECT @@IDENTITY", conn)
                             Dim insertedId = Convert.ToInt32(idCmd.ExecuteScalar())
                             ' Insert items
                             For Each row As DataGridViewRow In dgvInvoiceItems.Rows
                                 If row.IsNewRow Then Continue For
                                 Dim item = Convert.ToString(row.Cells("Description").Value)
                                 Dim qtyD As Double = 0D
                                 Dim upD As Double = 0D
                                 Double.TryParse(Convert.ToString(row.Cells("Qty").Value), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, qtyD)
                                 Double.TryParse(Convert.ToString(row.Cells("UnitPrice").Value), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, upD)
                                 Dim amountD As Double = qtyD * upD
                                 Dim itemCmd As New OleDbCommand("INSERT INTO InvoiceItems (InvoiceID, ItemDescription, Qty, UnitPrice, Amount) VALUES (?, ?, ?, ?, ?)", conn)
                                 itemCmd.Parameters.Add("p1", OleDbType.Integer).Value = insertedId
                                 itemCmd.Parameters.Add("p2", OleDbType.VarWChar).Value = item
                                 itemCmd.Parameters.Add("p3", OleDbType.Double).Value = qtyD
                                 itemCmd.Parameters.Add("p4", OleDbType.Double).Value = upD
                                 itemCmd.Parameters.Add("p5", OleDbType.Double).Value = amountD
                                 itemCmd.ExecuteNonQuery()
                             Next
                             Me.Invoke(Sub()
                                           MessageBox.Show("Invoice saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                           ' refresh serial for next invoice
                                           SetNewInvoiceSerial()
                                       End Sub)
                         End Using
                     Catch ex As Exception
                         Me.Invoke(Sub() MessageBox.Show("Failed to save invoice: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error))
                     End Try
                 End Sub)
    End Sub

    ' View history -> open FrmInvoiceHistory (modal)
    Private Sub btnViewHistory_Click(sender As Object, e As EventArgs)
        Try
            Dim f As New FrmInvoiceHistory(Me)
            f.ShowDialog()
        Catch ex As Exception
            MessageBox.Show("Failed to open history: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Load invoice for edit (called from history)
    Public Sub LoadInvoiceForEdit(id As Integer)
        If Not LicenseManager.IsLicenseValid(clientId) Then
            MessageBox.Show("Editing invoices requires a valid license.", "License Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Task.Run(Sub()
                     Try
                         Dim connStr = GetConnectionString()
                         Using conn As New OleDbConnection(connStr)
                             conn.Open()
                             Dim cmd As New OleDbCommand("SELECT InvoiceSerial, InvoiceDate, Client, Total FROM Invoices WHERE ID = ?", conn)
                             cmd.Parameters.Add("p1", OleDbType.Integer).Value = id
                             Using reader = cmd.ExecuteReader()
                                 If reader.Read() Then
                                     ' Read values safely to avoid invalid cast exceptions
                                     Dim serial As String = If(reader.IsDBNull(0), String.Empty, Convert.ToString(reader.GetValue(0)))
                                     Dim dt As DateTime = If(reader.IsDBNull(1), DateTime.Now, Convert.ToDateTime(reader.GetValue(1)))
                                     Dim client As String = If(reader.IsDBNull(2), String.Empty, Convert.ToString(reader.GetValue(2)))
                                     Dim total As Double = 0D
                                     If Not reader.IsDBNull(3) Then
                                         total = Convert.ToDouble(reader.GetValue(3))
                                     End If

                                     Me.Invoke(Sub()
                                                   txtInvoiceSerial.Text = serial
                                                   dtpInvoiceDate.Value = dt
                                                   txtBilledTo.Text = client
                                                   txtTotalCost.Text = total.ToString("N2")
                                                   dgvInvoiceItems.Rows.Clear()
                                               End Sub)

                                     ' Load items separately and safely
                                     Dim itemCmd As New OleDbCommand("SELECT ItemDescription, Qty, UnitPrice FROM InvoiceItems WHERE InvoiceID = ?", conn)
                                     itemCmd.Parameters.Add("p1", OleDbType.Integer).Value = id
                                     Using ista = itemCmd.ExecuteReader()
                                         While ista.Read()
                                             Dim desc As String = If(ista.IsDBNull(0), String.Empty, Convert.ToString(ista.GetValue(0)))
                                             Dim qty As Double = 0D
                                             Dim up As Double = 0D
                                             If Not ista.IsDBNull(1) Then qty = Convert.ToDouble(ista.GetValue(1))
                                             If Not ista.IsDBNull(2) Then up = Convert.ToDouble(ista.GetValue(2))
                                             Dim amount = qty * up
                                             Me.Invoke(Sub()
                                                           dgvInvoiceItems.Rows.Add("", desc, qty.ToString(), up.ToString("N2"), amount.ToString("N2"))
                                                       End Sub)
                                         End While
                                     End Using

                                     Me.Invoke(Sub()
                                                   editingInvoiceId = id
                                                   If pnlEditInvoice IsNot Nothing Then pnlEditInvoice.Visible = True
                                               End Sub)
                                 End If
                             End Using
                         End Using
                     Catch ex As Exception
                         Me.Invoke(Sub() MessageBox.Show("Failed to load invoice: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error))
                     End Try
                 End Sub)
    End Sub

    ' Save changes to the selected invoice
    Private Sub btnSaveChanges_Click(sender As Object, e As EventArgs)
        If editingInvoiceId <= 0 Then
            MessageBox.Show("No invoice selected for editing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If Not LicenseManager.IsLicenseValid(clientId) Then
            MessageBox.Show("Editing invoices requires a valid license.", "License Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Task.Run(Sub()
                     Try
                         Dim connStr = GetConnectionString()
                         Using conn As New OleDbConnection(connStr)
                             conn.Open()
                             ' basic header update (client, date, total)
                             Dim updateCmd As New OleDbCommand("UPDATE Invoices SET InvoiceSerial = ?, InvoiceDate = ?, Client = ?, Total = ? WHERE ID = ?", conn)
                             updateCmd.Parameters.Add("p1", OleDbType.VarWChar).Value = txtInvoiceSerial.Text
                             updateCmd.Parameters.Add("p2", OleDbType.Date).Value = dtpInvoiceDate.Value
                             updateCmd.Parameters.Add("p3", OleDbType.VarWChar).Value = txtBilledTo.Text
                             Dim tot As Double = 0D
                             Double.TryParse(txtTotalCost.Text, NumberStyles.Number, CultureInfo.CurrentCulture, tot)
                             updateCmd.Parameters.Add("p4", OleDbType.Double).Value = tot
                             updateCmd.Parameters.Add("p5", OleDbType.Integer).Value = editingInvoiceId
                             updateCmd.ExecuteNonQuery()
                             ' For simplicity, delete existing items and reinsert current grid items
                             Dim delCmd As New OleDbCommand("DELETE FROM InvoiceItems WHERE InvoiceID = ?", conn)
                             delCmd.Parameters.Add("p1", OleDbType.Integer).Value = editingInvoiceId
                             delCmd.ExecuteNonQuery()
                             For Each row As DataGridViewRow In dgvInvoiceItems.Rows
                                 If row.IsNewRow Then Continue For
                                 Dim item = Convert.ToString(row.Cells("Description").Value)
                                 Dim qtyD As Double = 0D
                                 Dim upD As Double = 0D
                                 Double.TryParse(Convert.ToString(row.Cells("Qty").Value), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, qtyD)
                                 Double.TryParse(Convert.ToString(row.Cells("UnitPrice").Value), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, upD)
                                 Dim itemCmd As New OleDbCommand("INSERT INTO InvoiceItems (InvoiceID, ItemDescription, Qty, UnitPrice, Amount) VALUES (?, ?, ?, ?, ?)", conn)
                                 itemCmd.Parameters.Add("p1", OleDbType.Integer).Value = editingInvoiceId
                                 itemCmd.Parameters.Add("p2", OleDbType.VarWChar).Value = item
                                 itemCmd.Parameters.Add("p3", OleDbType.Double).Value = qtyD
                                 itemCmd.Parameters.Add("p4", OleDbType.Double).Value = upD
                                 itemCmd.Parameters.Add("p5", OleDbType.Double).Value = qtyD * upD
                                 itemCmd.ExecuteNonQuery()
                             Next
                             Me.Invoke(Sub()
                                           MessageBox.Show("Invoice updated.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                           pnlEditInvoice.Visible = False
                                           editingInvoiceId = -1
                                       End Sub)
                         End Using
                     Catch ex As Exception
                         Me.Invoke(Sub() MessageBox.Show("Failed to update invoice: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error))
                     End Try
                 End Sub)
    End Sub
End Class

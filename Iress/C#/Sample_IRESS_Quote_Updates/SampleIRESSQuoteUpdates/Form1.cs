// © Copyright IRESS Limited. All Rights Reserved.
//
// IRESS Limited disclaims liability (including for negligence) for loss arising through access or use of this
// product or any information contained in it. IRESS Limited does not make any warranties of any kind in
// respect of the product or any information contained in it.
// 
// Summary:
// This sample provides a demonstration on how to use the IRESS Server API Type Library to access quote information 
// using the 'PricingQuoteGet' method.
// 
// Author: Simon Peckitt
// 
// Change History:
// 25-Feb-2014 [Simon Peckitt] Version 1.00: Released.
// 
// Dependencies:
// This sample requires IRESS Pro Neo 1.09 or higher and was built using Visual Studio 2012.
// 

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Diagnostics;

namespace SampleIRESSQuoteUpdates
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class Form1 : System.Windows.Forms.Form
    {
        private System.Windows.Forms.ListView PriceEventsListView;
        private System.Windows.Forms.ColumnHeader RequestedSecurityListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader SecurityListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader ExchangeListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader DataSourceListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader LastPriceListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader BidPriceListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader AskPriceListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader VolumeListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader OpenPriceListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader HighPriceListViewColumnHeader;
        private System.Windows.Forms.ColumnHeader LowPriceListViewColumnHeader;
        private System.Windows.Forms.Label SecurityLabel;
        private System.Windows.Forms.Label ExchangeLabel;
        private System.Windows.Forms.Label DataSourceLabel;
        private System.Windows.Forms.TextBox SecurityCodeTextBox;
        private System.Windows.Forms.TextBox ExchangeTextBox;
        private System.Windows.Forms.TextBox DataSourceTextBox;
        private System.Windows.Forms.Button AddButton;
        private System.Windows.Forms.Button RunButton;
        private System.Windows.Forms.Button StopButton;
        private System.Windows.Forms.Button ClearButton;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        // Form specific variables and constants.
        const string APPLICATION_TITLE = "Sample IRESS Quote Updates";
        private IressServerApi.RequestManager mRequestManager;
        private IressServerApi.Request mPricingQuoteGetRequest;
        private string[] msaOutputColumns = {"SecurityCode", "Exchange", "DataSource", "LastPrice", "BidPrice", "AskPrice", 
                                            "TotalVolume", "OpenPrice", "HighPrice", "LowPrice", "ErrorNumber"};
        private ArrayList maSecurities = new ArrayList();
        private ArrayList maExchanges = new ArrayList();
        private ArrayList maDataSources = new ArrayList();

        // This delegate enables asynchronous calls for updating the lstPriceEvents control with the initial quote snapshot.
        delegate void DisplayInitalPricesCallback(System.Array aData);

        // This delegate enables asynchronous calls for updating the lstPriceEvents control with a quote update.
        delegate void UpdatePricesCallback(System.Array aExtractedUpdateData);

        public Form1()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.PriceEventsListView = new System.Windows.Forms.ListView();
            this.RequestedSecurityListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SecurityListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.ExchangeListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.DataSourceListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.LastPriceListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.BidPriceListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.AskPriceListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.VolumeListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.OpenPriceListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.HighPriceListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.LowPriceListViewColumnHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SecurityLabel = new System.Windows.Forms.Label();
            this.ExchangeLabel = new System.Windows.Forms.Label();
            this.DataSourceLabel = new System.Windows.Forms.Label();
            this.SecurityCodeTextBox = new System.Windows.Forms.TextBox();
            this.ExchangeTextBox = new System.Windows.Forms.TextBox();
            this.DataSourceTextBox = new System.Windows.Forms.TextBox();
            this.AddButton = new System.Windows.Forms.Button();
            this.RunButton = new System.Windows.Forms.Button();
            this.StopButton = new System.Windows.Forms.Button();
            this.ClearButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // PriceEventsListView
            // 
            this.PriceEventsListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PriceEventsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.RequestedSecurityListViewColumnHeader,
            this.SecurityListViewColumnHeader,
            this.ExchangeListViewColumnHeader,
            this.DataSourceListViewColumnHeader,
            this.LastPriceListViewColumnHeader,
            this.BidPriceListViewColumnHeader,
            this.AskPriceListViewColumnHeader,
            this.VolumeListViewColumnHeader,
            this.OpenPriceListViewColumnHeader,
            this.HighPriceListViewColumnHeader,
            this.LowPriceListViewColumnHeader});
            this.PriceEventsListView.FullRowSelect = true;
            this.PriceEventsListView.GridLines = true;
            this.PriceEventsListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.PriceEventsListView.Location = new System.Drawing.Point(11, 130);
            this.PriceEventsListView.MultiSelect = false;
            this.PriceEventsListView.Name = "PriceEventsListView";
            this.PriceEventsListView.Size = new System.Drawing.Size(728, 174);
            this.PriceEventsListView.TabIndex = 9;
            this.PriceEventsListView.UseCompatibleStateImageBehavior = false;
            this.PriceEventsListView.View = System.Windows.Forms.View.Details;
            // 
            // RequestedSecurityListViewColumnHeader
            // 
            this.RequestedSecurityListViewColumnHeader.Text = "Requested Security";
            this.RequestedSecurityListViewColumnHeader.Width = 108;
            // 
            // SecurityListViewColumnHeader
            // 
            this.SecurityListViewColumnHeader.Text = "Security";
            // 
            // ExchangeListViewColumnHeader
            // 
            this.ExchangeListViewColumnHeader.Text = "Exch";
            // 
            // DataSourceListViewColumnHeader
            // 
            this.DataSourceListViewColumnHeader.Text = "DS";
            // 
            // LastPriceListViewColumnHeader
            // 
            this.LastPriceListViewColumnHeader.Text = "Last";
            this.LastPriceListViewColumnHeader.Width = 76;
            // 
            // BidPriceListViewColumnHeader
            // 
            this.BidPriceListViewColumnHeader.Text = "Bid";
            // 
            // AskPriceListViewColumnHeader
            // 
            this.AskPriceListViewColumnHeader.Text = "Ask";
            // 
            // VolumeListViewColumnHeader
            // 
            this.VolumeListViewColumnHeader.Text = "Volume";
            // 
            // OpenPriceListViewColumnHeader
            // 
            this.OpenPriceListViewColumnHeader.Text = "Open";
            // 
            // HighPriceListViewColumnHeader
            // 
            this.HighPriceListViewColumnHeader.Text = "High";
            // 
            // LowPriceListViewColumnHeader
            // 
            this.LowPriceListViewColumnHeader.Text = "Low";
            // 
            // SecurityLabel
            // 
            this.SecurityLabel.Location = new System.Drawing.Point(8, 16);
            this.SecurityLabel.Name = "SecurityLabel";
            this.SecurityLabel.Size = new System.Drawing.Size(64, 16);
            this.SecurityLabel.TabIndex = 0;
            this.SecurityLabel.Text = "Security:";
            // 
            // ExchangeLabel
            // 
            this.ExchangeLabel.Location = new System.Drawing.Point(8, 48);
            this.ExchangeLabel.Name = "ExchangeLabel";
            this.ExchangeLabel.Size = new System.Drawing.Size(64, 16);
            this.ExchangeLabel.TabIndex = 1;
            this.ExchangeLabel.Text = "Exchange";
            // 
            // DataSourceLabel
            // 
            this.DataSourceLabel.AutoSize = true;
            this.DataSourceLabel.Location = new System.Drawing.Point(8, 80);
            this.DataSourceLabel.Name = "DataSourceLabel";
            this.DataSourceLabel.Size = new System.Drawing.Size(67, 13);
            this.DataSourceLabel.TabIndex = 2;
            this.DataSourceLabel.Text = "Data Source";
            // 
            // SecurityCodeTextBox
            // 
            this.SecurityCodeTextBox.Location = new System.Drawing.Point(79, 15);
            this.SecurityCodeTextBox.Name = "SecurityCodeTextBox";
            this.SecurityCodeTextBox.Size = new System.Drawing.Size(96, 20);
            this.SecurityCodeTextBox.TabIndex = 3;
            // 
            // ExchangeTextBox
            // 
            this.ExchangeTextBox.Location = new System.Drawing.Point(79, 47);
            this.ExchangeTextBox.Name = "ExchangeTextBox";
            this.ExchangeTextBox.Size = new System.Drawing.Size(96, 20);
            this.ExchangeTextBox.TabIndex = 4;
            // 
            // DataSourceTextBox
            // 
            this.DataSourceTextBox.Location = new System.Drawing.Point(79, 79);
            this.DataSourceTextBox.Name = "DataSourceTextBox";
            this.DataSourceTextBox.Size = new System.Drawing.Size(96, 20);
            this.DataSourceTextBox.TabIndex = 5;
            // 
            // AddButton
            // 
            this.AddButton.Location = new System.Drawing.Point(200, 13);
            this.AddButton.Name = "AddButton";
            this.AddButton.Size = new System.Drawing.Size(104, 24);
            this.AddButton.TabIndex = 6;
            this.AddButton.Text = "Add";
            this.AddButton.Click += new System.EventHandler(this.AddButton_Click);
            // 
            // RunButton
            // 
            this.RunButton.Location = new System.Drawing.Point(350, 13);
            this.RunButton.Name = "RunButton";
            this.RunButton.Size = new System.Drawing.Size(104, 24);
            this.RunButton.TabIndex = 7;
            this.RunButton.Text = "Run";
            this.RunButton.Click += new System.EventHandler(this.RunButton_Click);
            // 
            // StopButton
            // 
            this.StopButton.Location = new System.Drawing.Point(500, 13);
            this.StopButton.Name = "StopButton";
            this.StopButton.Size = new System.Drawing.Size(104, 24);
            this.StopButton.TabIndex = 8;
            this.StopButton.Text = "Stop";
            this.StopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // ClearButton
            // 
            this.ClearButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.ClearButton.Location = new System.Drawing.Point(350, 77);
            this.ClearButton.Name = "ClearButton";
            this.ClearButton.Size = new System.Drawing.Size(104, 24);
            this.ClearButton.TabIndex = 10;
            this.ClearButton.Text = "Clear";
            this.ClearButton.Click += new System.EventHandler(this.ClearButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(751, 314);
            this.Controls.Add(this.PriceEventsListView);
            this.Controls.Add(this.SecurityLabel);
            this.Controls.Add(this.ExchangeLabel);
            this.Controls.Add(this.DataSourceLabel);
            this.Controls.Add(this.SecurityCodeTextBox);
            this.Controls.Add(this.ExchangeTextBox);
            this.Controls.Add(this.DataSourceTextBox);
            this.Controls.Add(this.AddButton);
            this.Controls.Add(this.RunButton);
            this.Controls.Add(this.StopButton);
            this.Controls.Add(this.ClearButton);
            this.Name = "Form1";
            this.Text = "Sample IRESS Quote Updates";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new Form1());
        }

        // Function:     AddButton
        // Purpose:      To add the security details entered to the PriceEventsListView control.
        // Assumptions:  The user will enter a valid security code at minimum with exchange and data source being optional.
        // Inputs:       [SecurityCodeTextBox] The security code, e.g BHP or VOD or BB. This is the only mandatory field.
        //               [ExchangeTextBox] The exchange specific to the security code entered.
        //               This value is not mandatory as the IRESS Server API can guess the prefered exchange based on your
        //               location information of the IRESS Pro configuration. If no Exchange is entered, it is expected 
        //               that no Data Source value would be entered.
        //               [DataSourceTextBox] The data source specific to the security code and exchange entered.
        //               This value is not mandatory as the IRESS Server API can guess the data source based on your
        //               location information of the IRESS Pro configuration.
        // Returns:      The values entered are added to the "Requested Security" column in the price quotes table.
        //               If any errors occur while adding the security details an exception error is thrown.
        private void AddButton_Click(object sender, System.EventArgs e)
        {
            try
            {
                String strSecurity, strExchange, strDataSource;
                strSecurity = SecurityCodeTextBox.Text.Trim();
                strExchange = ExchangeTextBox.Text.Trim();
                strDataSource = DataSourceTextBox.Text.Trim();

                if (String.Compare(strSecurity, String.Empty) == 0)
                {
                    MessageBox.Show("You must populate the security field first!");
                }
                else
                {
                    maSecurities.Add(strSecurity);

                    if (String.Compare(strExchange, String.Empty) == 0)
                    {
                        maExchanges.Add(System.DBNull.Value);
                    }
                    else
                    {
                        maExchanges.Add(strExchange);
                    }
                    if (String.Compare(strDataSource, String.Empty) == 0)
                    {
                        maDataSources.Add(System.DBNull.Value);
                    }
                    else
                    {
                        maDataSources.Add(strDataSource);
                    }

                    String strRequestedSecurity = strSecurity;

                    if (String.Compare(strExchange, String.Empty) != 0)
                    {
                        strRequestedSecurity = strRequestedSecurity + "." + strExchange;
                    }
                    if (String.Compare(strDataSource, String.Empty) != 0)
                    {
                        strRequestedSecurity = strRequestedSecurity + "@" + strDataSource;
                    }
                    ListViewItem lstItem = new ListViewItem(strRequestedSecurity);
                    lstItem.ForeColor = Color.Green;
                    lstItem.UseItemStyleForSubItems = false;
                    PriceEventsListView.Items.Add(lstItem);
                }
            }
            catch (Exception ex)
            {
                String strError = String.Format("Failed to add requested security to the quote table. Error: {0}", ex.Message);
                Debug.WriteLine(strError);
                DisplayError(strError);
            }
        }

        // Function:     DisplayError
        // Purpose:      To display an error to the user, using the error message specified.
        // Assumptions:  The error message is provided.
        // Inputs:       [ErrorMessage] The error message to display.
        // Returns:      A message box to the user with the error message.
        private void DisplayError(String ErrorMessage)
        {
            DialogResult result = MessageBox.Show(this, ErrorMessage, APPLICATION_TITLE, MessageBoxButtons.OK,
            MessageBoxIcon.Error);
        }

        // Function:     RunButton
        // Purpose:      To run the price quote request. This will then populate the price quotes table with retrieved data and updates from the 
        //               price quote request, for each security listed in the "Requested Security" column.
        // Assumptions:  The user will select the Run button after adding one or more securities to the price quotes table.
        // Inputs:       N/A.
        // Returns:      The price quote details for each security listed in the "Requested Security" column.
        private void RunButton_Click(object sender, System.EventArgs e)
        {
            AddButton.Enabled = false;
            RunButton.Enabled = false;
            RequestPriceQuotes();
        }

        // Function:     StopButton
        // Purpose:      To terminate the price quotes requests and any request updates, the price quotes table will stop updating.
        // Assumptions:  The user will select the Stop button after they have selected the Run button.
        // Inputs:       N/A
        // Returns:      N/A
        private void StopButton_Click(object sender, System.EventArgs e)
        {
            AddButton.Enabled = true;
            RunButton.Enabled = true;
            if (mPricingQuoteGetRequest != null)
            {
                mPricingQuoteGetRequest.StopWatchingUpdates();
                mPricingQuoteGetRequest = null;
            }
        }

        // Function:     ClearButton
        // Purpose:      To clear all the data from the table.
        // Assumptions:  The user will select the Clear button after they have added securities to the Requested Security list
        //               and/or run a price quote request.
        // Inputs:       N/A.
        // Returns:      The data in the price quotes table is cleared.
        private void ClearButton_Click(object sender, EventArgs e)
        {
            PriceEventsListView.Items.Clear();
            maSecurities.Clear();
            maExchanges.Clear();
            maDataSources.Clear();
        }

        // Function:     RequestPriceQuotes
        // Purpose:      To process the price quotes request, which retreives the price quote information for the securities listed in the 
        //               "Requested Security" column in the price quotes table.
        // Assumptions:  This function will run when the Run button is selected.
        // Inputs:       The security details from the "Requested Security" column in the price quotes table.
        // Returns:      The price quote data and updates for each requested security.
        //               If any errors occur during the price quote request an exception error is thrown.
        private void RequestPriceQuotes()
        {
            try
            {
                mRequestManager = new IressServerApi.RequestManager();

                mPricingQuoteGetRequest = mRequestManager.CreateMethod("IRESS", "", "PricingQuoteGet", IressServerApi.RequestDataType.HistoricalAndUpdates);

                mPricingQuoteGetRequest.DataReady += new IressServerApi.DataReadyDelegate(onDataReadyEvent);
                mPricingQuoteGetRequest.DataError += new IressServerApi.DataErrorDelegate(onDataErrorEvent);
                mPricingQuoteGetRequest.DataUpdate += new IressServerApi.DataUpdateDelegate(onDataUpdateEvent);

                mPricingQuoteGetRequest.Input.Header.Set("WaitForResponse", false, 0);

                for (int i = 0; i < maSecurities.Count; i++)
                {
                    mPricingQuoteGetRequest.Input.Parameters.Set("SecurityCode", maSecurities[i], i);
                    mPricingQuoteGetRequest.Input.Parameters.Set("Exchange", maExchanges[i], i);
                    mPricingQuoteGetRequest.Input.Parameters.Set("DataSource", maDataSources[i], i);
                }

                mPricingQuoteGetRequest.Execute();
            }

            catch (Exception ex)
            {
                String strError = String.Format("Failed to retrieve pricing quote information. Error: {0}", ex.Message);
                Debug.WriteLine(strError);
                DisplayError(strError);
            }
        }

        // Function:     onDataReadyEvent
        // Purpose:      Once the security data is ready and the data is retrieved for each security, display the results in the price quotes table.
        // Assumptions:  This function will run on a data ready event.
        // Inputs:       N/A
        // Returns:      The available price quote data for each security in the "Requested Security" column.
        private void onDataReadyEvent()
        {
            Debug.WriteLine("DataReadyEvent");
            System.Array aData;
            aData = (System.Array)mPricingQuoteGetRequest.Output.DataRows.GetRows(msaOutputColumns);
            DisplayData(aData);
        }

        // Function:     DisplayData
        // Purpose:      To populate the price quotes table with the price quote data retrieved.
        // Assumptions:  There is price quote data to display.
        // Inputs:       The price quote data retrieved from the price quotes request.
        // Returns:      The price quote data for each security listed in the "Requested Security" column from the price quotes table.
        //               If any errors occur, while displaying the data to the price quotes table, an exception error is thrown.
        private void DisplayData(System.Array aData)
        {
            try
            {
                if (this.PriceEventsListView.InvokeRequired)
                {
                    DisplayInitalPricesCallback d = new DisplayInitalPricesCallback(DisplayData);
                    this.Invoke(d, new object[] { aData });
                }
                else
                {
                    PriceEventsListView.Focus();
                    PriceEventsListView.Select();
                    double? dblLastPrice, dblBidPrice, dblAskPrice, dblTotalVolume, dblOpenPrice, dblHighPrice, dblLowPrice;
                    int nErrorNumber;

                    Int32 nIndex, nNumCodes = maSecurities.Count;
                    for (nIndex = 0; nIndex < nNumCodes; nIndex++)
                    {
                        dblLastPrice = (double?)aData.GetValue(nIndex, 3);
                        dblBidPrice = (double?)aData.GetValue(nIndex, 4);
                        dblAskPrice = (double?)aData.GetValue(nIndex, 5);
                        dblTotalVolume = (double?)aData.GetValue(nIndex, 6);
                        dblOpenPrice = (double?)aData.GetValue(nIndex, 7);
                        dblHighPrice = (double?)aData.GetValue(nIndex, 8);
                        dblLowPrice = (double?)aData.GetValue(nIndex, 9);
                        nErrorNumber = (int)aData.GetValue(nIndex, 10);

                        PriceEventsListView.Items[nIndex].SubItems.Add(aData.GetValue(nIndex, 0).ToString()); //Security
                        PriceEventsListView.Items[nIndex].SubItems.Add(aData.GetValue(nIndex, 1).ToString()); //Exchange
                        PriceEventsListView.Items[nIndex].SubItems.Add(aData.GetValue(nIndex, 2).ToString()); //Data Source
                        PriceEventsListView.Items[nIndex].SubItems.Add(dblLastPrice == null ? "" : dblLastPrice.ToString()); //Last Price
                        PriceEventsListView.Items[nIndex].SubItems.Add(dblBidPrice == null ? "" : dblBidPrice.ToString()); //Bid Price
                        PriceEventsListView.Items[nIndex].SubItems.Add(dblAskPrice == null ? "" : dblAskPrice.ToString()); //Ask Price
                        PriceEventsListView.Items[nIndex].SubItems.Add(dblTotalVolume == null ? "" : dblTotalVolume.ToString()); //Total Volume
                        PriceEventsListView.Items[nIndex].SubItems.Add(dblOpenPrice == null ? "" : dblOpenPrice.ToString()); //Open Price
                        PriceEventsListView.Items[nIndex].SubItems.Add(dblHighPrice == null ? "" : dblHighPrice.ToString()); //High Price
                        PriceEventsListView.Items[nIndex].SubItems.Add(dblLowPrice == null ? "" : dblLowPrice.ToString()); //Low Price

                        if (nErrorNumber != 0)
                        {
                            for (int i = 1; i <= 10; i++)
                                PriceEventsListView.Items[nIndex].SubItems[i].ForeColor = Color.Red;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                String strError = String.Format("Failed to output pricing information to the quote table. Error: {0}", ex.Message);
                Debug.WriteLine(strError);
                DisplayError(strError);
            }
        }

        // Function:     onDataErrorEvent
        // Purpose:      To catch any errors during the price quote request.
        // Assumptions:  N/A
        // Inputs:       N/A.
        // Returns:      A message box with the data error event message.
        private void onDataErrorEvent()
        {
            string strError = String.Format("An error has occurred retrieving the price quote information. Error: {0}",
                                mPricingQuoteGetRequest.Output.Error.Get("Message").ToString());
            Debug.WriteLine("DataErrorEvent " + strError);
            DisplayError(strError);
        }

        // Function:     onDataUpdateEvent
        // Purpose:      To catch price quote updates.
        // Assumptions:  This function will run automatically and continuously until the Stop button is selected by the user.
        // Inputs:       N/A
        // Returns:      The updated price quote data for the corresponding security listed in the "Requested Security" column from the price quotes table.
        private void onDataUpdateEvent()
        {
            Debug.WriteLine("DataUpdateEvent");

            // COM server has received some update data and it can be now be retrieved.
            // All the update data must be retrieved before the DataUpdate event will be raised again.
            Object aUpdateData;

            // Defaults to retrieving all rows (and deleting them from the COM server).
            aUpdateData = mPricingQuoteGetRequest.Output.UpdateRows.GetRowsAndRemove();

            // Getting update data values in bulk
            System.Array aExtractedUpdateData;
            aExtractedUpdateData = (System.Array)mPricingQuoteGetRequest.Output.UpdateRows.GetRowsFromRetrievedData(msaOutputColumns, ref aUpdateData);

            UpdateData(aExtractedUpdateData);
        }

        // Function:     UpdateData
        // Purpose:      To update the price quotes table as new data is received from an Update event. 
        // Assumptions:  This function is called when a data update event occurs.
        // Inputs:       N/A
        // Returns:      The updated price quote data to the price quotes table, for the corresponding security listed in the "Requested Security" column.
        private void UpdateData(System.Array aExtractedUpdateData)
        {
            try
            {
                if (this.PriceEventsListView.InvokeRequired)
                {
                    UpdatePricesCallback d = new UpdatePricesCallback(UpdateData);
                    this.Invoke(d, new object[] { aExtractedUpdateData });
                }
                else
                {
                    Int32 nIndex, nUpdateCount = aExtractedUpdateData.GetUpperBound(0) + 1;
                    for (nIndex = 0; nIndex < nUpdateCount; nIndex++)
                    {
                        string strSecurity = aExtractedUpdateData.GetValue(nIndex, 0).ToString();
                        string strExchange = aExtractedUpdateData.GetValue(nIndex, 1).ToString();
                        string strDataSource = aExtractedUpdateData.GetValue(nIndex, 2).ToString();

                        for (int nRow = 0; nRow < PriceEventsListView.Items.Count; nRow++)
                        {
                            if ((PriceEventsListView.Items[nRow].SubItems[1].Text == strSecurity) &&
                            (PriceEventsListView.Items[nRow].SubItems[2].Text == strExchange) &&
                            (PriceEventsListView.Items[nRow].SubItems[3].Text == strDataSource))
                            {
                                double? dblLastPrice, dblBidPrice, dblAskPrice, dblTotalVolume, dblOpenPrice, dblHighPrice, dblLowPrice;

                                dblLastPrice = (double?)aExtractedUpdateData.GetValue(nIndex, 3);
                                dblBidPrice = (double?)aExtractedUpdateData.GetValue(nIndex, 4);
                                dblAskPrice = (double?)aExtractedUpdateData.GetValue(nIndex, 5);
                                dblTotalVolume = (double?)aExtractedUpdateData.GetValue(nIndex, 6);
                                dblOpenPrice = (double?)aExtractedUpdateData.GetValue(nIndex, 7);
                                dblHighPrice = (double?)aExtractedUpdateData.GetValue(nIndex, 8);
                                dblLowPrice = (double?)aExtractedUpdateData.GetValue(nIndex, 9);

                                if (dblLastPrice != null)
                                    PriceEventsListView.Items[nRow].SubItems[4].Text = dblLastPrice.ToString();

                                if (dblBidPrice != null)
                                    PriceEventsListView.Items[nRow].SubItems[5].Text = dblBidPrice.ToString();

                                if (dblAskPrice != null)
                                    PriceEventsListView.Items[nRow].SubItems[6].Text = dblAskPrice.ToString();

                                if (dblTotalVolume != null)
                                    PriceEventsListView.Items[nRow].SubItems[7].Text = dblTotalVolume.ToString();

                                if (dblOpenPrice != null)
                                    PriceEventsListView.Items[nRow].SubItems[8].Text = dblOpenPrice.ToString();

                                if (dblHighPrice != null)
                                    PriceEventsListView.Items[nRow].SubItems[9].Text = dblHighPrice.ToString();

                                if (dblLowPrice != null)
                                    PriceEventsListView.Items[nRow].SubItems[10].Text = dblLowPrice.ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                String strError = String.Format("Failed to update pricing information in the quote table. Error: {0}", ex.Message);
                Debug.WriteLine(strError);
                DisplayError(strError);
            }
        }
    }
}

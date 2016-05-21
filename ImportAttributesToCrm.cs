using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using InformationPanel = XrmToolBox.Extensibility.InformationPanel;

namespace CRMConsultants.CreateAttributes
{
    public partial class ImportAttributesToCrm : PluginControlBase
    {
        #region Constructor
        public  const int _languageCode = 1033;
        public ImportAttributesToCrm()
        {
            InitializeComponent();
            initialize();
        }

        private void initialize()
        {
            //Attribute Type	Attribute Schema Name	Option Set Values	Attribute Display Name	Description	RequiredLevel	Boolean Default  Value	String Format	String Length	Date Type Format	Integer Format	Integer Minimum Value	Integer Maximum Value	Floating Number Precision	Float Min Value	Float Max Value	Decimal Precision	Decimal Min Value	Decimal Max Value	Currency precision	Currency Min Value	Currency Max Value	 IME Mode	AuditEnable	IsValidForAdvancedFind
            ApplicationSetting.ColumnsListToVarify= new string[]{ "Attribute Type", "Attribute Schema Name", "Option Set Values", "Attribute Display Name", "Description", "RequiredLevel",   "Boolean Default  Value",  "String Format",   "String Length",   "Date Type Format",   "Integer Format",  "Integer Minimum Value",   "Integer Maximum Value",   "Floating Number Precision",   "Float Min Value", "Float Max Value", "Decimal Precision",   "Decimal Min Value",   "Decimal Max Value",   "Currency precision",  "Currency Min Value",  "Currency Max Value",   "IME Mode",   "AuditEnable", "IsValidForAdvancedFind" };
        }

        #endregion Constructor

        private void tsbCloseThisTab_Click(object sender, EventArgs e)
        {
            CloseTool();
        }
        private void LoadEntities()
        {
            cmbEntities.Enabled = true;
            cmbEntities.Items.Clear();
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving Entities...",
                AsyncArgument = null,
                Work = (bw, e) =>
                {
                    var request = new RetrieveAllEntitiesRequest { EntityFilters = EntityFilters.Entity };
                    var response = (RetrieveAllEntitiesResponse)Service.Execute(request);

                    e.Result = response.EntityMetadata;
                },

                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        MessageBox.Show(this, "Error occured: " + e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        var emds = (EntityMetadata[])e.Result;

                        foreach (var emd in emds)
                        {
                            if (emd.IsCustomizable.Value==true && emd.CanCreateAttributes.Value==true)
                                cmbEntities.Items.Add(new EntityData(emd.LogicalName, emd.DisplayName != null && emd.DisplayName.UserLocalizedLabel != null ? emd.DisplayName.UserLocalizedLabel.Label : "N/A", emd.PrimaryIdAttribute));
                        }
                    }
                },
                ProgressChanged = e => { SetWorkingMessage(e.UserState.ToString()); }
            });
        }
        private void tsbDownLoadTemplate_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = "xls";
            saveFileDialog.Filter = "Excel files (*.xls)|*.xls |All files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                Assembly asm = Assembly.GetExecutingAssembly();
                string file = string.Format("{0}.Attribute Sample.xlsx", asm.GetName().Name);
                Stream fileStream = asm.GetManifestResourceStream(file);
                SaveStreamToFile(saveFileDialog.FileName, fileStream);  //<--here is where to save to disk

                this.Invoke(new Action(() => { MessageBox.Show(this, "File Downloaded successfully."); }));

            }
        }

        public void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            if (stream.Length == 0) return;

            // Create a FileStream object to write a stream to a file
            using (FileStream fileStream = System.IO.File.Create(fileFullPath, (int)stream.Length))
            {
                // Fill the bytes[] array with the stream data
                byte[] bytesInStream = new byte[stream.Length];
                stream.Read(bytesInStream, 0, (int)bytesInStream.Length);

                // Use FileStream object to write to the specified file
                fileStream.Write(bytesInStream, 0, bytesInStream.Length);
            }
        }

        private void btnBrowseFile_Click(object sender, EventArgs e)
        {
            if(cmbEntities.SelectedItem==null)
            {
                MessageBox.Show(this, "Please select Entity.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var dialog = new SaveFileDialog
            {
                Filter = "Excel workbook|*.xlsx",
                Title = "Select a location for the file generated"
            };
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                txtFilePath.Text = dialog.FileName;
                Assembly assembly = Assembly.GetExecutingAssembly();
                Assembly asm = Assembly.GetExecutingAssembly();
                string file = string.Format("{0}.Attribute Sample.xlsx", asm.GetName().Name);
                Stream fileStream = asm.GetManifestResourceStream(file);
                SaveStreamToFile(dialog.FileName, fileStream); 
                this.Invoke(new Action(() => { txtFilePath.Text = dialog.FileName; }));
                this.Invoke(new Action(() => { MessageBox.Show(this, "File Downloaded successfully.");
                    tsbCreateAttributes.Enabled = true;
                }));
            }
        }

        private void tsbCreateAttributes_Click(object sender, EventArgs e)
        {
            CreateExcelDoc excell_app = null;
            ApplicationSetting.ExcelData = new System.Data.DataTable();
            ApplicationSetting.FileToImport = "";
            int progressCounter = 0;
            var dialog = new OpenFileDialog
            {
                Filter = "Excel workbook|*.xlsx",
                Title = "Select a file to import"
            };

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                ApplicationSetting.FileToImport = dialog.FileName;
            }
            else
            {
                MessageBox.Show(this, "Please select file to import.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var responseAlert = MessageBox.Show(this, "Do you want to create Attributes?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (responseAlert != DialogResult.OK)
                return;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Importing Attributes...",
                AsyncArgument = null,
                Work = (bw, em) =>
                {

                    excell_app = new CreateExcelDoc(ApplicationSetting.FileToImport.Trim());
                    int rowCount = excell_app.workSheet_range.Rows.Count;
                    int colCount = excell_app.workSheet_range.Columns.Count;
                    int row = 0;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        bool addRow = true;
                        // Validate and Insert only defined values
                        int col = 0;
                        for (int j = 1; j <= colCount; j++)
                        {
                            if (excell_app.workSheet_range.Cells[i, j] != null && excell_app.workSheet_range.Cells[i, j].Value2 != null)
                            {
                                string val = excell_app.workSheet_range.Cells[i, j].Value2.ToString();
                                //Use the first row to add columns to DataTable.
                                if (i == 1)
                                {
                                    // if (ApplicationSetting.ColumnsListToVarify.Contains(val))
                                    ApplicationSetting.ExcelData.Columns.Add(val.Trim());
                                }
                                else
                                {
                                    //Add rows to DataTable.
                                    if (addRow)
                                    {
                                        ApplicationSetting.ExcelData.Rows.Add();
                                        addRow = false;
                                    }
                                    // if ((col + 1) <= ApplicationSetting.ExcelData.Columns.Count)
                                    ApplicationSetting.ExcelData.Rows[i - 2][col] = val.Trim();
                                }
                            }
                            //if ((col + 1) <= ApplicationSetting.ExcelData.Columns.Count)
                            col++;
                        }
                        row++;
                    }

                    if (!IsValidNoOfColumns())
                    {
                        this.Invoke(new Action(() =>
                        {
                            MessageBox.Show(this, "Columns do not match to template column list.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }));
                        return;
                    }
                    else if (!IsValidRequiredColumns())
                    {
                        this.Invoke(new Action(() =>
                        {
                            MessageBox.Show(this, "Please make sure Attribute Display Name, Attribute Schema Name and Attribute Type is entered.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }));
                        return;
                    }
                    // Loop through Dataset and Create Attributes
                    if (ApplicationSetting.ExcelData != null && ApplicationSetting.ExcelData.Rows.Count > 0)
                    {
                        foreach (DataRow rw in ApplicationSetting.ExcelData.Rows)
                        {
                            try
                            {
                                bw.ReportProgress(progressCounter * 100 / ApplicationSetting.ExcelData.Rows.Count, string.Concat("Importing Attributes..."));
                                CreateCRMAttribute(rw);
                            }
                            catch
                            {
                                // If any attribute creation got failed log it. instead of terminating
                            }
                            progressCounter++;
                        }
                    }
                    bw.ReportProgress(100, "Attributes created successfully........");
                    this.Invoke(new Action(() =>
                    {
                        MessageBox.Show(this,"Attributes created successfully........");
                    }));
                    this.Invoke(new Action(() =>
                    {
                        excell_app.ReleaseObject();
                        excell_app = null;
                        txtFilePath.Text = "";               
                    }));
                },

                PostWorkCallBack = em =>
                {
                    if (em.Error != null)
                    {
                        this.Invoke(new Action(() =>
                        {
                            MessageBox.Show(this, "Error occured: " + em.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                },
                ProgressChanged = em => { SetWorkingMessage(em.UserState.ToString()); }
            });
        }

        private void tsbLoadEntities_Click(object sender, EventArgs e)
        {
            LoadEntities();
        }

        private void cmbEntities_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbEntities.SelectedItem == null)
            {
                return;
            }
            btnBrowseFile.Enabled = true;
            ApplicationSetting.SelectedEntity = (EntityData)cmbEntities.SelectedItem;
        }

        public bool IsValidRequiredColumns()
        {
            if (ApplicationSetting.ExcelData.Rows.Count > 0)
            {
                var results = ApplicationSetting.ExcelData.Select("[Attribute Type] is null OR [Attribute Display Name] is null OR [Attribute Schema Name] is null");

                if (results != null && results.Count() > 0)
                    return false;
            }
            return true;
        }

        public bool IsValidNoOfColumns()
        {
            if (ApplicationSetting.ExcelData.Rows.Count > 0)
            {
                foreach (DataColumn column in ApplicationSetting.ExcelData.Columns)
                {
                    if (!ApplicationSetting.ColumnsListToVarify.Contains(column.ColumnName))
                        return false;
                }
            }
            return true;
        }

        public void CreateCRMAttribute(DataRow record)
        {
            string result = string.Empty;
            try
            {
                #region # Read From Datatable #
                AttributeMetadata createMetadata = new AttributeMetadata();
                bool isGlobal = false;
                AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                string reqLevelText = "";
                int precisionSource = 0;
                int currencyPrecision = 2;
                if (record["RequiredLevel"] != null && !string.IsNullOrEmpty(Convert.ToString(record["RequiredLevel"])))
                {
                    reqLevelText = Convert.ToString(record["RequiredLevel"]).ToLower();
                    requirementLevel = reqLevelText == "required" ? AttributeRequiredLevel.ApplicationRequired : AttributeRequiredLevel.Recommended;
                }
                reqLevelText = record["Attribute Type"].ToString().ToLower();
                string attributeSchemaName= record["Attribute Schema Name"].ToString().ToLower();
                string optionSetValues = record["Option Set Values"].ToString().ToLower();
                string attributeDisplayName = record["Attribute Display Name"].ToString().ToLower();
                string attributeDiscription = record["Description"].ToString().ToLower();
                bool boolDefaultValue= record["Boolean Default  Value"].ToString().ToLower()=="yes"?true:false;

                Microsoft.Xrm.Sdk.Metadata.StringFormat stringFormat = GetStringFormat(record["String Format"].ToString().ToLower());
               int stringLength= record["String Length"]!=null && !string.IsNullOrEmpty(Convert.ToString(record["String Length"])) && Convert.ToInt32(record["String Length"].ToString()) <=4000? Convert.ToInt32(record["String Length"].ToString()) : 4000;

                Microsoft.Xrm.Sdk.Metadata.DateTimeFormat dateFormat = record["Date Type Format"]!=null && !string.IsNullOrEmpty(record["Date Type Format"].ToString()) ? GetDateFormat(record["Date Type Format"].ToString().ToLower()):DateTimeFormat.DateAndTime;
                Microsoft.Xrm.Sdk.Metadata.IntegerFormat integerFormt = record["Integer Format"] != null && !string.IsNullOrEmpty(record["Integer Format"].ToString()) ? GetIntegerFormat(record["Integer Format"].ToString().ToLower()) : IntegerFormat.None;

                Double intMinValue = record["Integer Minimum Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Integer Minimum Value"])) && Convert.ToDouble(record["Integer Minimum Value"].ToString())<= 2147483647 && Convert.ToDouble(record["Integer Minimum Value"].ToString()) >= -2147483647 ? Convert.ToDouble(record["Integer Minimum Value"].ToString()) : -2147483648;
                Double intMaxValue = record["Integer Maximum Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Integer Maximum Value"])) &&  Convert.ToDouble(record["Integer Maximum Value"].ToString()) <= 2147483647 && Convert.ToDouble(record["Integer Maximum Value"].ToString()) >= -2147483647 ? Convert.ToDouble(record["Integer Maximum Value"].ToString()) : 2147483647;

                int floatingPrecision = record["Floating Number Precision"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Floating Number Precision"])) ? Convert.ToInt32(record["Floating Number Precision"].ToString()) : 2;
                Double floatMinValue = record["Float Min Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Float Min Value"])) && Convert.ToDouble(record["Float Min Value"].ToString()) >=0 && Convert.ToDouble(record["Float Min Value"].ToString()) <= 1000000000 ? Convert.ToDouble(record["Float Min Value"].ToString()) : 0;
                Double floatMaxValue = record["Float Max Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Float Max Value"])) && Convert.ToDouble(record["Float Max Value"].ToString()) >= 0 && Convert.ToDouble(record["Float Max Value"].ToString()) <= 1000000000 ? Convert.ToDouble(record["Float Max Value"].ToString()) : 1000000000;

                int decimalPrecision = record["Decimal Precision"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Decimal Precision"])) ? Convert.ToInt32(record["Decimal Precision"].ToString()) : 2;
                Decimal decimalMinValue = record["Decimal Min Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Decimal Min Value"])) && Convert.ToDecimal(record["Decimal Min Value"].ToString())>=-100000000000 && Convert.ToDecimal(record["Decimal Min Value"].ToString()) <= 100000000000 ? Convert.ToDecimal(record["Decimal Min Value"].ToString()) : -100000000000;
                Decimal decimalMaxValue = record["Decimal Max Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Decimal Max Value"])) && Convert.ToDecimal(record["Decimal Max Value"].ToString()) >= -100000000000 && Convert.ToDecimal(record["Decimal Max Value"].ToString()) <= 100000000000 ? Convert.ToDecimal(record["Decimal Max Value"].ToString()) : 100000000000;

                Double currencyMinValue = record["Currency Min Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Currency Min Value"])) && Convert.ToDouble(record["Currency Min Value"].ToString())>=-922337203685477 && Convert.ToDouble(record["Currency Min Value"].ToString())<= 922337203685477 ? Convert.ToDouble(record["Currency Min Value"].ToString()) : -922337203685477;
                Double currencyMaxValue = record["Currency Max Value"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Currency Max Value"])) && Convert.ToDouble(record["Currency Max Value"].ToString()) >= -922337203685477 && Convert.ToDouble(record["Currency Max Value"].ToString()) <= 922337203685477 ? Convert.ToDouble(record["Currency Max Value"].ToString()) : 922337203685477;

                Microsoft.Xrm.Sdk.Metadata.ImeMode imeMode = GetIMEMode(record["IME Mode"].ToString().ToLower());

                if(record["Currency precision"] != null && !string.IsNullOrEmpty(Convert.ToString(record["Currency precision"]).ToLower()))
                {
                    switch(Convert.ToString(record["Currency precision"]).ToLower().Trim())
                    {
                        case "pricing decimal precision":
                            precisionSource = 1;
                            break;
                        case "currency precision":
                            precisionSource = 2;
                            break;
                        default:
                            currencyPrecision = Convert.ToInt32(Convert.ToString(record["Currency precision"]));
                            break;
                    }
                }

                bool isAuditEnabled = record["AuditEnable"]!=null && record["AuditEnable"].ToString().ToLower() == "yes" ? true : false;
                bool isValidForAdvancFind = record["IsValidForAdvancedFind"]!=null && record["IsValidForAdvancedFind"].ToString().ToLower() == "yes" ? true : false;
                #endregion # Read From Datatable #

                switch (reqLevelText.ToLower().Trim())
                {
                    case "boolean":
                        // Create a boolean attribute
                        createMetadata = new BooleanAttributeMetadata
                        {
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            // Set extended properties
                            OptionSet = new BooleanOptionSetMetadata(
                                new OptionMetadata(new Microsoft.Xrm.Sdk.Label("Yes", _languageCode), 1),
                                new OptionMetadata(new Microsoft.Xrm.Sdk.Label("No", _languageCode), 0)
                                ),
                            DefaultValue=boolDefaultValue,
                            IsAuditEnabled = new BooleanManagedProperty(isAuditEnabled),
                            IsValidForAdvancedFind = new BooleanManagedProperty(isValidForAdvancFind),                            
                        };
                        break;
                    case "date and time":
                        createMetadata = new DateTimeAttributeMetadata
                        {
                            // Set base properties
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            // Set extended properties
                            Format = dateFormat,
                            ImeMode = imeMode,
                            IsAuditEnabled = new BooleanManagedProperty(isAuditEnabled),
                            IsValidForAdvancedFind = new BooleanManagedProperty(isValidForAdvancFind),                                    
                        };
                        break;
                    case "multiple line of text":
                        createMetadata = new MemoAttributeMetadata
                        {
                            // Set base properties
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            // Set extended properties
                            ImeMode = imeMode,
                            IsAuditEnabled = new BooleanManagedProperty(isAuditEnabled),
                            IsValidForAdvancedFind = new BooleanManagedProperty(isValidForAdvancFind),
                            MaxLength = stringLength
                        };
                        break;
                    case "whole number":
                        createMetadata = new IntegerAttributeMetadata
                        {
                            // Set base properties
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            // Set extended properties
                           // ImeMode = imeMode,// in crm 2016 ths feature is there
                            // Set extended properties
                            Format = IntegerFormat.None,
                            MaxValue =Convert.ToInt32( intMaxValue),
                            MinValue = Convert.ToInt32(intMinValue),
                            IsAuditEnabled = new BooleanManagedProperty(isAuditEnabled),
                            IsValidForAdvancedFind = new BooleanManagedProperty(isValidForAdvancFind)
                        };
                        break;
                    case "floating point number":
                        createMetadata = new DoubleAttributeMetadata
                        {
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            MaxValue = floatMaxValue,
                            MinValue = floatMinValue,
                            Precision=floatingPrecision,
                            IsAuditEnabled = new BooleanManagedProperty(isAuditEnabled),
                            IsValidForAdvancedFind = new BooleanManagedProperty(isValidForAdvancFind),
                            ImeMode = imeMode
                        };
                        break;
                    case "decimal number":
                        createMetadata = new DecimalAttributeMetadata
                        {
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            MaxValue = decimalMaxValue,
                            MinValue = decimalMinValue,
                            Precision = decimalPrecision,
                            IsAuditEnabled = new BooleanManagedProperty(isAuditEnabled),
                            IsValidForAdvancedFind = new BooleanManagedProperty(isValidForAdvancFind),
                            ImeMode = imeMode
                        };
                        break;
                    case "currency":
                        createMetadata = new MoneyAttributeMetadata
                        {
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            MaxValue = currencyMaxValue,
                            MinValue = currencyMinValue,
                            Precision = currencyPrecision,
                            PrecisionSource = precisionSource,
                            ImeMode = imeMode
                        };
                        break;
                    case "option set":

                        OptionMetadataCollection optionMetadataCollection = GetOptionMetadata(optionSetValues);
                        OptionSetMetadata Optionmedata = new OptionSetMetadata();
                        if(optionMetadataCollection!=null && optionMetadataCollection.Count()>0)
                           Optionmedata.Options.AddRange(optionMetadataCollection);
                        Optionmedata.IsGlobal = isGlobal;
                        Optionmedata.OptionSetType = OptionSetType.Picklist;

                        createMetadata = new PicklistAttributeMetadata
                        {
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            OptionSet = Optionmedata
                        };
                        break;
                    case "single line of text":
                        createMetadata = new StringAttributeMetadata
                        {
                            SchemaName = attributeSchemaName,
                            DisplayName = new Microsoft.Xrm.Sdk.Label(attributeDisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                            Description = new Microsoft.Xrm.Sdk.Label(attributeDiscription, _languageCode),
                            // Set extended properties
                            ImeMode = imeMode,
                            IsAuditEnabled = new BooleanManagedProperty(isAuditEnabled),
                            IsValidForAdvancedFind = new BooleanManagedProperty(isValidForAdvancFind),
                            MaxLength = stringLength
                        };
                        break;
                }

                //ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                //{
                //    // Assign settings that define execution behavior: continue on error, return responses. 
                //    Settings = new ExecuteMultipleSettings()
                //    {
                //        ContinueOnError = false,
                //        ReturnResponses = true
                //    },
                //    // Create an empty organization request collection.
                //    Requests = new OrganizationRequestCollection()
                //};
                CreateAttributeRequest request = new CreateAttributeRequest
                {
                    Attribute = createMetadata,
                    EntityName = ApplicationSetting.SelectedEntity.LogicalName
                };
                try
                {
                    Service.Execute(request);
                    result = "Success";
                }
                catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
                {
                    result = ex.Message;
                }
                catch (Exception ex)
                {
                    result = ex.Message;
                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }            
        }

        private static Microsoft.Xrm.Sdk.Metadata.ImeMode GetIMEMode(string imeMode)
        {
            // In acse of none we need min value & max value
            Microsoft.Xrm.Sdk.Metadata.ImeMode returnImeMode= Microsoft.Xrm.Sdk.Metadata.ImeMode.Auto;
            switch (imeMode)
            {
                case "active":
                    returnImeMode = Microsoft.Xrm.Sdk.Metadata.ImeMode.Active;
                    break;

                case "inactive":
                    returnImeMode = Microsoft.Xrm.Sdk.Metadata.ImeMode.Inactive ;
                    break;

                case "disabled":
                    returnImeMode = Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled;
                    break;

            }
            return returnImeMode;
        }

        private static IntegerFormat GetIntegerFormat(string format)
        {
            // In acse of none we need min value & max value
            Microsoft.Xrm.Sdk.Metadata.IntegerFormat returnFormat = Microsoft.Xrm.Sdk.Metadata.IntegerFormat.None;
            switch (format)
            {
                case "duration":
                    returnFormat = Microsoft.Xrm.Sdk.Metadata.IntegerFormat.Duration ;
                    break;

                case "time zone":
                    returnFormat = Microsoft.Xrm.Sdk.Metadata.IntegerFormat.TimeZone;
                    break;

                case "language":
                    returnFormat = Microsoft.Xrm.Sdk.Metadata.IntegerFormat.Language;
                    break;
            }
            return returnFormat;
        }

        private static DateTimeFormat GetDateFormat(string dateFormat)
        {
            Microsoft.Xrm.Sdk.Metadata.DateTimeFormat format = Microsoft.Xrm.Sdk.Metadata.DateTimeFormat.DateAndTime;
            switch (dateFormat)
            {
                case "date only":
                    format = Microsoft.Xrm.Sdk.Metadata.DateTimeFormat.DateOnly;
                    break;

                case "date and time":
                    format = Microsoft.Xrm.Sdk.Metadata.DateTimeFormat.DateAndTime;
                    break;
            }
            return format;
        }

        private static Microsoft.Xrm.Sdk.Metadata.StringFormat GetStringFormat(string formatText)
        {
            Microsoft.Xrm.Sdk.Metadata.StringFormat format = Microsoft.Xrm.Sdk.Metadata.StringFormat.Text;
            switch (formatText)
            {
                case "email":
                    format=Microsoft.Xrm.Sdk.Metadata.StringFormat.Email;
                    break;

                case "text":
                    format= Microsoft.Xrm.Sdk.Metadata.StringFormat.Text;
                    break;

                case "textarea":
                    format= Microsoft.Xrm.Sdk.Metadata.StringFormat.TextArea;
                    break;

                case "url":
                    format= Microsoft.Xrm.Sdk.Metadata.StringFormat.Url;
                    break;

                case "ticker symbol":
                    format= Microsoft.Xrm.Sdk.Metadata.StringFormat.TickerSymbol;
                    break;

                case "phone":
                    format= Microsoft.Xrm.Sdk.Metadata.StringFormat.Phone;
                    break;
            }
            return format;
        }
        private static Microsoft.Xrm.Sdk.Metadata.OptionMetadataCollection GetOptionMetadata(string option)
        {
            OptionMetadataCollection optionMetadataCollection = new OptionMetadataCollection();
            if (option != "")
            {
                if (option.Contains("|"))
                {
                    string[] optionArray = option.Split('|');
                    if (optionArray != null && optionArray.Length > 0)
                    {
                        for (int arrayCounter = 0; arrayCounter < optionArray.Length; arrayCounter++)
                        {
                            optionMetadataCollection.Add(new OptionMetadata(
                                new Microsoft.Xrm.Sdk.Label(optionArray[arrayCounter], _languageCode), null));
                        }
                    }
                }
                else
                {
                    optionMetadataCollection.Add(new OptionMetadata(
                               new Microsoft.Xrm.Sdk.Label(option, _languageCode), null));
                }
            }
            return optionMetadataCollection;
        }
        //private static void AddBooleanAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue)
        //{
        //    BooleanAttributeMetadata booleanAttributeMetadata = new BooleanAttributeMetadata();
        //    booleanAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    booleanAttributeMetadata.SchemaName = attributeName;
        //    booleanAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    booleanAttributeMetadata.OptionSet = new BooleanOptionSetMetadata(new OptionMetadata(new Microsoft.Xrm.Sdk.Label("True", 1033), new int?(1)), new OptionMetadata(new Microsoft.Xrm.Sdk.Label("False", 1033), new int?(0)));
        //    CreateMetadata._newAttribute.addedAttributes = booleanAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

        //private static void AddStringAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue, int maxLength)
        //{
        //    StringAttributeMetadata stringAttributeMetadata = new StringAttributeMetadata();
        //    stringAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    stringAttributeMetadata.SchemaName = attributeName;
        //    stringAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    stringAttributeMetadata.MaxLength = new int?(maxLength);
        //    CreateMetadata._newAttribute.addedAttributes = stringAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

        //private static void AddMoneyAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue)
        //{
        //    MoneyAttributeMetadata moneyAttributeMetadata = new MoneyAttributeMetadata();
        //    moneyAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    moneyAttributeMetadata.SchemaName = attributeName;
        //    moneyAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    moneyAttributeMetadata.MaxValue = new double?(1000000000.0);
        //    moneyAttributeMetadata.MinValue = new double?(0.0);
        //    moneyAttributeMetadata.Precision = new int?(2);
        //    moneyAttributeMetadata.PrecisionSource = new int?(2);
        //    moneyAttributeMetadata.ImeMode = new Microsoft.Xrm.Sdk.Metadata.ImeMode?(Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled);
        //    CreateMetadata._newAttribute.addedAttributes = moneyAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

        //private static void AddDecimalAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue)
        //{
        //    DecimalAttributeMetadata decimalAttributeMetadata = new DecimalAttributeMetadata();
        //    decimalAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    decimalAttributeMetadata.SchemaName = attributeName;
        //    decimalAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    decimalAttributeMetadata.MaxValue = new decimal?(1000000000m);
        //    decimalAttributeMetadata.MinValue = new decimal?(0m);
        //    decimalAttributeMetadata.Precision = new int?(2);
        //    decimalAttributeMetadata.ImeMode = new Microsoft.Xrm.Sdk.Metadata.ImeMode?(Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled);
        //    CreateMetadata._newAttribute.addedAttributes = decimalAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

        //private static void AddMemoAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue, int maxLength)
        //{
        //    MemoAttributeMetadata memoAttributeMetadata = new MemoAttributeMetadata();
        //    memoAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    memoAttributeMetadata.SchemaName = attributeName;
        //    memoAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    memoAttributeMetadata.Format = new StringFormat?(StringFormat.TextArea);
        //    memoAttributeMetadata.ImeMode = new Microsoft.Xrm.Sdk.Metadata.ImeMode?(Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled);
        //    memoAttributeMetadata.MaxLength = new int?(maxLength);
        //    CreateMetadata._newAttribute.addedAttributes = memoAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

        //private static void AddIntegerAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue, int maxLength)
        //{
        //    IntegerAttributeMetadata integerAttributeMetadata = new IntegerAttributeMetadata();
        //    integerAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    integerAttributeMetadata.SchemaName = attributeName;
        //    integerAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    integerAttributeMetadata.Format = new IntegerFormat?(IntegerFormat.None);
        //    integerAttributeMetadata.MinValue = new int?(0);
        //    integerAttributeMetadata.MaxValue = new int?(maxLength);
        //    CreateMetadata._newAttribute.addedAttributes = integerAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

        //private static void AddDateAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue)
        //{
        //    DateTimeAttributeMetadata dateTimeAttributeMetadata = new DateTimeAttributeMetadata();
        //    dateTimeAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    dateTimeAttributeMetadata.SchemaName = attributeName;
        //    dateTimeAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    dateTimeAttributeMetadata.Format = new DateTimeFormat?(DateTimeFormat.DateOnly);
        //    dateTimeAttributeMetadata.ImeMode = new Microsoft.Xrm.Sdk.Metadata.ImeMode?(Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled);
        //    CreateMetadata._newAttribute.addedAttributes = dateTimeAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

        //private static void AddOptionSetAttribute(string entityName, string attributeName, AttributeRequiredLevel requirementLevel, string displayValue, bool isGlobal, OptionSetType type)
        //{
        //    PicklistAttributeMetadata picklistAttributeMetadata = new PicklistAttributeMetadata();
        //    picklistAttributeMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(displayValue, 1033);
        //    picklistAttributeMetadata.SchemaName = attributeName;
        //    picklistAttributeMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel);
        //    picklistAttributeMetadata.OptionSet = new OptionSetMetadata
        //    {
        //        IsGlobal = new bool?(isGlobal),
        //        OptionSetType = new OptionSetType?(type)
        //    };
        //    CreateMetadata._newAttribute.addedAttributes = picklistAttributeMetadata;
        //    CreateMetadata._newAttribute.attributeEntity = entityName;
        //    CreateMetadata.addedAttributes.Add(CreateMetadata._newAttribute);
        //}

    }
}
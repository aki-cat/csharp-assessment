using Parquet.Schema;
using Parquet;
using System.Data;
using System.Text;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.Axes;

namespace ArcVera_Tech_Test
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private async void btnImportEra5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Parquet files (*.parquet)|*.parquet|All files (*.*)|*.*";
                openFileDialog.Title = "Select a Parquet File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    DataTable dataTable = await ReadParquetFileAsync(filePath);
                    dgImportedEra5.DataSource = dataTable;
                    PlotU10DailyValues(dataTable);
                }
            }
        }

        private async Task<DataTable> ReadParquetFileAsync(string filePath)
        {
            using (Stream fileStream = File.OpenRead(filePath))
            {
                using (var parquetReader = await ParquetReader.CreateAsync(fileStream))
                {
                    DataTable dataTable = new DataTable();

                    for (int i = 0; i < parquetReader.RowGroupCount; i++)
                    {
                        using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                        {
                            // Create columns
                            foreach (DataField field in parquetReader.Schema.GetDataFields())
                            {
                                if (!dataTable.Columns.Contains(field.Name))
                                {
                                    Type columnType = field.HasNulls ? typeof(object) : field.ClrType;
                                    dataTable.Columns.Add(field.Name, columnType);
                                }

                                // Read values from Parquet column
                                DataColumn column = dataTable.Columns[field.Name];
                                Array values = (await groupReader.ReadColumnAsync(field)).Data;
                                for (int j = 0; j < values.Length; j++)
                                {
                                    if (dataTable.Rows.Count <= j)
                                    {
                                        dataTable.Rows.Add(dataTable.NewRow());
                                    }
                                    dataTable.Rows[j][field.Name] = values.GetValue(j);
                                }
                            }
                        }
                    }

                    return dataTable;
                }
            }
        }

        private void PlotU10DailyValues(DataTable dataTable)
        {
            var plotModel = new PlotModel { Title = "Daily u10 Values" };
            var lineSeries = new LineSeries { Title = "u10" };

            var groupedData = dataTable.AsEnumerable()
                .GroupBy(row => DateTime.Parse(row["date"].ToString()))
                .Select(g => new
                {
                    Date = g.Key,
                    U10Average = g.Average(row => Convert.ToDouble(row["u10"]))
                })
                .OrderBy(data => data.Date);

            foreach (var data in groupedData)
            {
                lineSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(data.Date), data.U10Average));
            }

            plotModel.Series.Add(lineSeries);
            plotView1.Model = plotModel;
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            ExportCsv();
        }

        private void ExportCsv()
        {
            if (dgImportedEra5.DataSource is not DataTable dataTable)
            {
                MessageBox.Show("No data to export. Import data first.");
                return;
            }

            string raw = SerializeToCsv(dataTable);
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog.Title = "Save CSV File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (Stream fileStream = File.OpenWrite(saveFileDialog.FileName))
                    {
                        fileStream.Write(Encoding.UTF8.GetBytes(raw), 0, raw.Length);
                        fileStream.Close();
                    }
                }
            }
        }

        private string SerializeToCsv(DataTable dataTable)
        {
            StringBuilder builder = new();

            string[] rowBuffer = new string[dataTable.Columns.Count];
            for (int i = 0; i < rowBuffer.Length; i++)
            {
                DataColumn col = dataTable.Columns[i];
                rowBuffer[i] = col.ColumnName;
            }

            builder.AppendJoin(",", rowBuffer);
            builder.Append("\n");

            foreach (DataRow row in dataTable.Rows)
            {
                for (int i = 0; i < rowBuffer.Length; i++)
                {
                    rowBuffer[i] = row[i].ToString() ?? "";
                }

                builder.AppendJoin(",", rowBuffer);
                builder.Append("\n");
            }

            return builder.ToString();
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            ExportExcel();
        }

        private void ExportExcel()
        {
            if (dgImportedEra5.DataSource is not DataTable dataTable)
            {
                MessageBox.Show("No data to export. Import data first.");
                return;
            }

            string raw = SerializeToExcel(dataTable);
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel files (*.xlsx)|(*.xlsx)|All files (*.*)|*.*";
                saveFileDialog.Title = "Save Excel File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (Stream fileStream = File.OpenWrite(saveFileDialog.FileName))
                    {
                        fileStream.Write(Encoding.UTF8.GetBytes(raw), 0, raw.Length);
                        fileStream.Close();
                    }
                }
            }
        }

        private string SerializeToExcel(DataTable dataTable)
        {
            return "Test empty excel file.";
        }
    }
}

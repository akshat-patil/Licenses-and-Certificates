using System;
using System.Data.SqlClient;
using ClosedXML.Excel;

namespace EmsConsole
{
    public class SqlData
    {
        private readonly string _connectionString;

        public SqlData(string connectionString)
        {
            _connectionString = connectionString;
        }

        public void ProcessFile(string excelFilePath)
        {
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);

                // --- Find GatewayName & DeviceName ---
                string gatewayName = FindSingleValue(worksheet, "Gateway Name");
                string deviceName = FindSingleValue(worksheet, "Device Name");

                if (string.IsNullOrWhiteSpace(gatewayName) || string.IsNullOrWhiteSpace(deviceName))
                {
                    Console.WriteLine($"Skipping {excelFilePath}, missing Gateway/Device");
                    return;
                }

                int deviceId = InsertDevice(gatewayName, deviceName);

                // --- Find multiple value columns ---
                int tsCol = FindColumn(worksheet, "Local Time Stamp");
                int energyCol = FindColumn(worksheet, "TotalDeliveredActiveEnergy (Wh)", "TotalActiveDeliveredEnergy (Wh)", "Active energy delivered (Wh)");
                int errorCol = FindColumn(worksheet, "Error");

                if (tsCol == -1 || energyCol == -1 || errorCol == -1)
                {
                    Console.WriteLine($"Skipping {excelFilePath}, missing required data columns.");
                    return;
                }

                int startRow = FindHeaderRow(worksheet, "Local Time Stamp"); // header row index
                int row = startRow + 1;

                using (var conn = new SqlConnection(_connectionString))
                {
                    conn.Open();

                    while (!worksheet.Cell(row, tsCol).IsEmpty())
                    {
                        string timestamp = worksheet.Cell(row, tsCol).GetValue<string>();
                        string energy = worksheet.Cell(row, energyCol).GetValue<string>();
                        string error = worksheet.Cell(row, errorCol).GetValue<string>();

                        InsertEnergyReading(conn, deviceId, timestamp, energy, error);
                        row++;
                    }
                }
            }
        }

        // Find a single-value header (e.g., Gateway Name, Device Name) and return value below it
        private string FindSingleValue(IXLWorksheet ws, string header)
        {
            foreach (var cell in ws.CellsUsed())
            {
                if (cell.GetValue<string>().Trim().Equals(header, StringComparison.OrdinalIgnoreCase))
                {
                    return cell.Worksheet.Cell(cell.Address.RowNumber + 1, cell.Address.ColumnNumber).GetValue<string>().Trim();
                }
            }
            return null;
        }

        // Find the row index of a header
        private int FindHeaderRow(IXLWorksheet ws, string header)
        {
            foreach (var cell in ws.CellsUsed())
            {
                if (cell.GetValue<string>().Trim().Equals(header, StringComparison.OrdinalIgnoreCase))
                    return cell.Address.RowNumber;
            }
            return -1;
        }

        // Find column index of a header
        // Updated FindColumn method that supports multiple possible headers
        private int FindColumn(IXLWorksheet ws, params string[] headers)
        {
            foreach (var cell in ws.CellsUsed())
            {
                string cellValue = cell.GetValue<string>().Trim();
                foreach (var header in headers)
                {
                    if (cellValue.Equals(header, StringComparison.OrdinalIgnoreCase))
                        return cell.Address.ColumnNumber;
                }
            }
            return -1;
        }


        // Insert device if not exists, return DeviceId
        private int InsertDevice(string gateway, string device)
        {
            using (var conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    IF NOT EXISTS (SELECT 1 FROM ems.Device WHERE GatewayName=@g AND DeviceName=@d)
                    BEGIN
                        INSERT INTO ems.Device(GatewayName, DeviceName) VALUES(@g,@d);
                    END
                    SELECT DeviceId FROM ems.Device WHERE GatewayName=@g AND DeviceName=@d;";

                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@g", gateway);
                    cmd.Parameters.AddWithValue("@d", device);
                    return (int)cmd.ExecuteScalar();
                }
            }
        }

        // Insert readings into EnergyReadings
        private void InsertEnergyReading(SqlConnection conn, int deviceId, string ts, string energy, string error)
        {
            string sql = @"INSERT INTO ems.EnergyReadings (DeviceId, LocalTimestamp, ActiveEnergyDelivered, Error) 
                           VALUES (@id,@ts,@energy,@err)";

            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@id", deviceId);
                cmd.Parameters.AddWithValue("@ts", ts ?? "");
                cmd.Parameters.AddWithValue("@energy", energy ?? "");
                cmd.Parameters.AddWithValue("@err", error ?? "");
                cmd.ExecuteNonQuery();
            }
        }
    }
}

# Đường dẫn đến file Excel và tên sheet chứa dữ liệu SQL
$excel_file = "./SQL_Server_2016_Audit.xlsx"
$sheet_name = "Sheet1"

# Import module ImportExcel nếu chưa được import
if (-not (Get-Module -Name ImportExcel -ErrorAction SilentlyContinue)) {
    Import-Module -Name ImportExcel
}

# Đọc dữ liệu từ file Excel
$data = Import-Excel -Path $excel_file -WorksheetName $sheet_name

# Duyệt qua từng hàng trong dữ liệu
foreach ($row in $data) {
    # Lấy lệnh SQL từ cột "Audit"
    $sql_command = $row.Audit
    # Lấy số thứ tự cột để ghi kết quả
    $result_column = $row.Result

    # Kiểm tra xem lệnh SQL có tồn tại không
    if ($sql_command -ne "") {
        # Thực thi lệnh SQL và nhận kết quả
        $result = Invoke-SqlCmd -ServerInstance 'QUANGVINH\MSSQLSERVER02' -Database 'master' -Query $sql_command
        
        # Chuyển kết quả thành chuỗi
        $resultString = $result | Out-String

        # Ghi kết quả vào cột "Result"
        $row."Result" = $resultString
    }
}

# Ghi dữ liệu đã được cập nhật vào file Excel
$data | Export-Excel -Path $excel_file -WorksheetName $sheet_name -AutoSize -ClearSheet -Show

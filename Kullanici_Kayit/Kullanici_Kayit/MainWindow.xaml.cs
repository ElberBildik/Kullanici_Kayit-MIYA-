using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Müsteri_Aktarımı
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ExcelYukleButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog.ShowDialog() == true)
            {
                string fileName = openFileDialog.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(fileName);
                Excel.Worksheet worksheet = workbook.Worksheets[1];

                DataTable dt = new DataTable();

                // Sütun başlıklarını DataTable'a ekleme
                int colCount = worksheet.UsedRange.Columns.Count;
                for (int col = 1; col <= colCount; col++)
                {
                    Excel.Range headerCell = worksheet.Cells[1, col];
                    dt.Columns.Add(headerCell.Value2 != null ? headerCell.Value2.ToString() : "");
                }

                // Veri satırlarını DataTable'a ekleme
                int rowCount = worksheet.UsedRange.Rows.Count;
                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow newRow = dt.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        Excel.Range cell = worksheet.Cells[row, col];
                        newRow[col - 1] = cell.Value2 != null ? cell.Value2.ToString() : ".";
                    }
                    dt.Rows.Add(newRow);
                }

                // DataTable'ı DataGrid'e bağlama
                dataGridView.ItemsSource = dt.DefaultView;

                // Excel işlemlerini kapatma
                workbook.Close();
                excelApp.Quit();

               

                // Kullanıcıya bilgi vermek için mesaj kutusu göster
                MessageBox.Show("Excel dosyası başarıyla yüklendi ", "Bilgi", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async void AktarimiBaslatButton_Click(object sender, RoutedEventArgs e)
        {
            string panServisLinki = textBoxPanServisLink.Text;
            string panServisSifresi = textBoxPanServisSifresi.Text;
            string dist = textBoxDist.Text;
            string firmaKodu = textBoxFirmaKodu.Text;
            string calismaYili = textBoxCalismaYili.Text;

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {panServisSifresi}");

                    var values = new Dictionary<string, string>
            {
                { "dist", dist },
                { "firmaKodu", firmaKodu },
                { "calismaYili", calismaYili }
            };

                    var content = new FormUrlEncodedContent(values);
                    HttpResponseMessage response = await client.GetAsync(panServisLinki);

                    if (response.IsSuccessStatusCode)
                    {
                        // API yanıtını okuma
                        string responseString = await response.Content.ReadAsStringAsync();

                        // Yanıtı DataGrid'e yazdırma (örnek olarak)
                        DataTable dt = new DataTable();
                        dt.Columns.Add("Response");
                        dt.Rows.Add(responseString);
                        dataGridView.ItemsSource = dt.DefaultView;

                        // Başarı mesajı gösterme
                        MessageBox.Show($"Başarılı: {responseString}", "Başarı", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // Hata durumunu işleme
                        string errorString = await response.Content.ReadAsStringAsync();
                        MessageBox.Show($"Hata: {response.StatusCode}\nMesaj: {errorString}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                // Genel istisna durumunu işleme
                MessageBox.Show($"İstek gönderilirken bir hata oluştu: {ex.Message}", "İstisna", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

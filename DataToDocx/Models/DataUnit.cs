using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Wpf.Ui.Controls;
using System.IO;
using System.Windows.Data;

namespace DataToDocx.Models
{
    public class DataUnit:ObservableObject
    {
        public string tableName { get; set; }
        public string? key { get; set; }

        public List<string>? attrs { get; set; }
        public string? FilePath { get; set; }
        private ICommand fileChoose;
        public ICommand FileChoose
        {
            get
            {
                fileChoose = new RelayCommand(ChooseFile);
                return fileChoose;
            }
            set { fileChoose = value; }
        }

        private ICommand fileUpload;
        public ICommand FileUpload
        {
            get {
                fileUpload = new RelayCommand(Loadfile);
                return fileUpload; }
            set { fileUpload = value; }
        }
        private ICommand delSelf;

        public ICommand DelSelf
        {
            get
            {
                delSelf = new RelayCommand(DelSelfMessage);
                return delSelf;
            }
            set { delSelf = value; }
        }

        private string connstr;

        public string Connstr
        {
            get {
                connstr ??= AppConfig.SqliteConnstr;
                return connstr; }
            set { connstr = value; }
        }

        public DataUnit(string tablename)
        {
            tableName = tablename;
        }

        public void  ChooseFile()
        {
            //string[] fileName = (string[])e.Data.GetData(DataFormats.FileDrop);
            //FilePath = fileName[0];
            //OnPropertyChanged(nameof(FilePath));
            OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel file（*.xlsx）|*.xlsx|ALL files（*.*）|*.*",
                InitialDirectory = this.FilePath
            };
            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                OnPropertyChanged(nameof(FilePath));
            }

            
        }

        public void DelSelfMessage()
        {
            WeakReferenceMessenger.Default.Send<string, string>(tableName, "DelDataUnits");
        }

        public void Loadfile()
        {
            if (FilePath is null || FilePath == "")
            {
                return;
            }
            if (Path.GetExtension(FilePath) != ".xlsx")
            {
                //WeakReferenceMessenger.Default.Send<Snackbar>(new Snackbar()
                //{
                //    Title = "文件格式错误！",
                //    Message = "导入文件格式不是xlsx,请重新选择文件。",
                //    Icon = Wpf.Ui.Common.SymbolRegular.ErrorCircle24,
                //    Appearance = Wpf.Ui.Common.ControlAppearance.Caution
                //});
                FilePath = "";
                OnPropertyChanged(nameof(FilePath));
                return;
            }

            //InputTask(FilePath, AppConfig.SqliteCnn, TabName);
        }




    }
}

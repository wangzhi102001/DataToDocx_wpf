// This Source Code Form is subject to the terms of the MIT License.
// If a copy of the MIT was not distributed with this file, You can obtain one at https://opensource.org/licenses/MIT.
// Copyright (C) Leszek Pomianowski and WPF UI Contributors.
// All Rights Reserved.

using CommunityToolkit.Mvvm.Messaging;
using DataToDocx.Models;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SQLite;
using System.Windows.Media;
using System.Windows.Threading;
using Wpf.Ui.Controls;
using Microsoft.Data.Sqlite;
using System.IO;


namespace DataToDocx.ViewModels.Pages
{
    public partial class TemplateViewModel : ObservableObject, INavigationAware
    {
        private bool _isInitialized = false;
        private ObservableCollection<DataTab> datatabs;

        public ObservableCollection<DataTab> DataTabs
        {
            get { return datatabs; }
            set { datatabs = value; OnPropertyChanged(nameof(DataTabs)); }
        }
        [ObservableProperty]
        private int dBCmbSelectIndex = 0;


        [ObservableProperty]
        private Database database;
        private bool _isDelEnd = true;
        public bool IsDelEnd
        {
            get { return _isDelEnd; }
            set { _isDelEnd = value; OnPropertyChanged(nameof(IsDelEnd)); }
        }
        private ObservableCollection<Database> dataBases;
        public ObservableCollection<Database> DataBases
        {
            get
            {
                //Dbload();
                IsDelEnd = false;
                OnPropertyChanged(nameof(IsDelEnd));
                dataBases ??= AppConfig.Dbload();
                IsDelEnd = true;
                OnPropertyChanged(nameof(IsDelEnd));

                return dataBases;
            }
            set
            {
                dataBases = value;
                OnPropertyChanged(nameof(DataBases));
                //WeakReferenceMessenger.Default.Send(DataBases);
            }
        }

        public void OnNavigatedTo()
        {
            if (!_isInitialized)
                InitializeViewModel();

            new Task(() =>
            {
                IsDelEnd = false;
                OnPropertyChanged(nameof(IsDelEnd));
                DataBases = AppConfig.Dbload();
                OnPropertyChanged(nameof(DataBases));
                IsDelEnd = true;
                OnPropertyChanged(nameof(IsDelEnd));

            }).Start();
        }

        public void OnNavigatedFrom() { }

        private void InitializeViewModel()
        {
            

            _isInitialized = true;
        }

        [RelayCommand]
        public void OnDelDB()
        {

            IsDelEnd = false;
            OnPropertyChanged(nameof(IsDelEnd));
            if (DBCmbSelectIndex != -1)
            {
                try
                {
                    SqliteConnection.ClearAllPools();
                    File.Delete(DataBases[DBCmbSelectIndex].Path);

                }
                catch (Exception io)
                {
                    Fun.ShowSnackbar( "删除失败", $"错误提示：{io.Message}",3);

                    Fun.Updatelogtext($"数据库删除失败,错误提示：{io.Message}");
                    IsDelEnd = true;
                    OnPropertyChanged(nameof(IsDelEnd));
                    return;
                }

                Fun.ShowSnackbar($"删除成功。",$"数据库【{DataBases[DBCmbSelectIndex].Name}】已删除。"
                    ,4);
                Fun.Updatelogtext($"数据库【{DataBases[DBCmbSelectIndex].Name}】已删除。");


                DataBases.RemoveAt(DBCmbSelectIndex);
                OnPropertyChanged(nameof(DataBases));
                IsDelEnd = true;
                OnPropertyChanged(nameof(IsDelEnd));
            }
            else
            {
                Fun.ShowSnackbar("删除失败", $"没有数据库被选中" ,3);
                Fun.Updatelogtext("数据库删除失败,没有数据库被选中");

            }
            IsDelEnd = true;
            OnPropertyChanged(nameof(IsDelEnd));



        }
    }
}

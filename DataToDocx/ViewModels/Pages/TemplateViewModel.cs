// This Source Code Form is subject to the terms of the MIT License.
// If a copy of the MIT was not distributed with this file, You can obtain one at https://opensource.org/licenses/MIT.
// Copyright (C) Leszek Pomianowski and WPF UI Contributors.
// All Rights Reserved.

using DataToDocx.Models;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows.Media;
using Wpf.Ui.Controls;

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
    }
}

// This Source Code Form is subject to the terms of the MIT License.
// If a copy of the MIT was not distributed with this file, You can obtain one at https://opensource.org/licenses/MIT.
// Copyright (C) Leszek Pomianowski and WPF UI Contributors.
// All Rights Reserved.

using CommunityToolkit.Mvvm.Messaging;
using DataToDocx.Models;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using Wpf.Ui.Controls;

namespace DataToDocx.ViewModels.Pages
{
    public partial class DashboardViewModel : ObservableObject
    {
        private ObservableCollection<DataUnit> dataUnits = new ObservableCollection<DataUnit>();
        public ObservableCollection<DataUnit> DataUnits { 
            
            get {  return dataUnits; }
        
        set { dataUnits = value; OnPropertyChanged(nameof(DataUnits)); }
        }

        private ObservableCollection<DataUnit> maindataUnits = new ObservableCollection<DataUnit>();
        public ObservableCollection<DataUnit> MainDataUnits
        {

            get { return maindataUnits; }

            set { maindataUnits = value; OnPropertyChanged(nameof(MainDataUnits)); }
        }

        private ISnackbarService _snackbarService;
        private IContentDialogService _contentDialogService;

        public DashboardViewModel(ISnackbarService snackbarService,IContentDialogService contentDialogService)
        {
            _snackbarService = snackbarService;
            _contentDialogService = contentDialogService;
            WeakReferenceMessenger.Default.Register<List<string>,string>(this,"snackbarError",ShowSnackBarError);
            WeakReferenceMessenger.Default.Register<List<string>, string>(this, "snackbarSuccess", ShowSnackBarSuccess);
            WeakReferenceMessenger.Default.Register<List<string>, string>(this, "snackbarCaution", ShowSnackBarCaution);
            WeakReferenceMessenger.Default.Register<List<string>, string>(this, "snackbarInfo", ShowSnackBarInfo);
            WeakReferenceMessenger.Default.Register<List<string>, string>(this, "dialogAlart", ShowDialogAlart);
            WeakReferenceMessenger.Default.Register<string, string>(this, "DelDataUnits", (r, message) =>
            {
                foreach (var item in DataUnits.ToList())
                {
                    if (item.tableName==message)
                    {
                        DataUnits.Remove(item);
                    }
                }
            });
            MainDataUnits = new ObservableCollection<DataUnit>()
            {
                new DataUnit("主数据库")
            };

            //WeakReferenceMessenger.Default.Register<SimpleContentDialogCreateOptions, string>(this, "dialogContent", ShowDialogContent);


        }
        private void ShowSnackBarError(object recipient, List<string> strings)
        {

            _snackbarService.Show(strings[0], strings[1],ControlAppearance.Danger,new SymbolIcon(Wpf.Ui.Common.SymbolRegular.Info24));
        }
        private void ShowSnackBarSuccess(object recipient, List<string> strings)
        {

            _snackbarService.Show(strings[0], strings[1], ControlAppearance.Success, new SymbolIcon(Wpf.Ui.Common.SymbolRegular.Info24));
        }
        private void ShowSnackBarCaution(object recipient, List<string> strings)
        {

            _snackbarService.Show(strings[0], strings[1], ControlAppearance.Caution, new SymbolIcon(Wpf.Ui.Common.SymbolRegular.Info24));
        }
        private void ShowSnackBarInfo(object recipient, List<string> strings)
        {

            _snackbarService.Show(strings[0], strings[1], ControlAppearance.Info, new SymbolIcon(Wpf.Ui.Common.SymbolRegular.Info24));
        }

        private void ShowDialogAlart(object recipient, List<string> strings)
        {
            _contentDialogService.ShowAlertAsync(strings[0], strings[1], strings[2]);
        }

        [RelayCommand]
        private void AddSecondryData()
        {
            int i = 1;
            if (DataUnits.Count==0)
            {
                DataUnits.Add(new DataUnit("DataTable1"));
            }
            else
            {
                foreach (var item in DataUnits.ToList())
                {
                    if (item.tableName==$"DataTable{i}")
                    {
                        i++;
                    }
                }
                DataUnits.Add(new DataUnit($"DataTable{i}"));
            }            
            OnPropertyChanged(nameof(DataUnits));
        }



    }
}

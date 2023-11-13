// This Source Code Form is subject to the terms of the MIT License.
// If a copy of the MIT was not distributed with this file, You can obtain one at https://opensource.org/licenses/MIT.
// Copyright (C) Leszek Pomianowski and WPF UI Contributors.
// All Rights Reserved.

using DataToDocx.ViewModels.Pages;
using Wpf.Ui.Controls;

namespace DataToDocx.Views.Pages
{
    public partial class CreatePage : INavigableView<CreateViewModel>
    {
        public CreateViewModel ViewModel { get; }

        public CreatePage(CreateViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
        }
    }
}

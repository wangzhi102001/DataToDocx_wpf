// This Source Code Form is subject to the terms of the MIT License.
// If a copy of the MIT was not distributed with this file, You can obtain one at https://opensource.org/licenses/MIT.
// Copyright (C) Leszek Pomianowski and WPF UI Contributors.
// All Rights Reserved.

using CommunityToolkit.Mvvm.Messaging;
using DataToDocx.Models;
using System.Windows.Documents;
using System.Windows.Media;
using Wpf.Ui.Controls;

namespace DataToDocx.ViewModels.Pages
{
    public partial class LogViewModel : ObservableObject, INavigationAware
    {
        private bool _isInitialized = false;

        

        public void OnNavigatedTo()
        {
            if (!_isInitialized)
                InitializeViewModel();
        }

        public void OnNavigatedFrom() { }


        public LogViewModel() {

            WeakReferenceMessenger.Default.Register<string,string>(this,"log", OnReceive_plus);
        }

        private void InitializeViewModel()
        {
            //WeakReferenceMessenger.Default.Register<string>(this, OnReceive_plus);

            _isInitialized = true;
        }

        private FlowDocument _document = new()
        {
            FontFamily = new System.Windows.Media.FontFamily("Microsoft YaHei"),
            LineHeight = 5,
        };

        public FlowDocument Document
        {
            get { return _document; }
            set { _document = value; OnPropertyChanged(nameof(Document)); }
        }
        private void OnReceive_plus(object recipient, string message)
        {
            if (!message.StartsWith("20"))
            {
                AppendToRTB(message);
            }
            else
            {
                AppendToRTB($"{message}");
            }
        }
        private void AppendToRTB(string text)
        {
            //Application.Current.Dispatcher.BeginInvoke

            Application.Current.Dispatcher.BeginInvoke((Action)(() =>
            {
                Paragraph p = new Paragraph();
                Run run = new Run(text);
                p.Inlines.Add(run);


                if (Document.Blocks.FirstBlock != null)
                {
                    if (Document.Blocks.Count <= 1000)
                    {
                        Document.Blocks.InsertBefore(Document.Blocks.FirstBlock, p);
                    }
                    else
                    {
                        Document.Blocks.Remove(Document.Blocks.LastBlock);//防止日志页卡顿。限制日志页条数为1000条。
                        Document.Blocks.InsertBefore(Document.Blocks.FirstBlock, p);
                    }
                }
                else
                {
                    Document.Blocks.Add(p);
                }
                OnPropertyChanged(nameof(Document));

            }));


            //Run r = new Run(text);
            //Paragraph para = new Paragraph();
            //para.Inlines.Add(r);
            //RichTextBox.Document.Blocks.Clear();
            //RichTextBox.Document.Blocks.Add(para);
            //OnPropertyChanged(nameof(RichTextBox));

        }
    }
}

﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Word
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.MinWidth = 250;
            this.MinHeight = 400;
        }

        private void Open_Word_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Create_Word_Click(object sender, RoutedEventArgs e)
        {
            Redactor redactor = new Redactor();
            redactor.Show();
            this.Close();
        }

        private void Open_Excel_Click(object sender, RoutedEventArgs e)
        {
            Redactor4 redactor = new Redactor4();
            redactor.Show();
            this.Close();
        }

        private void Create_Excel_Click(object sender, RoutedEventArgs e)
        {
            Redactor2 redactor = new Redactor2();
            redactor.Show();
            this.Close();
        }
    }
}

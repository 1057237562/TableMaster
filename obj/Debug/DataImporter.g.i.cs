﻿#pragma checksum "..\..\DataImporter.xaml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "F323BBB4EC02E66A7533E3DFDE557C20B8D54083"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using MahApps.Metro.Controls;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace TableMaster {
    
    
    /// <summary>
    /// DataImporter
    /// </summary>
    public partial class DataImporter : MahApps.Metro.Controls.MetroWindow, System.Windows.Markup.IComponentConnector {
        
        
        #line 12 "..\..\DataImporter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Path;
        
        #line default
        #line hidden
        
        
        #line 13 "..\..\DataImporter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button Browser;
        
        #line default
        #line hidden
        
        
        #line 14 "..\..\DataImporter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TabControl TabControl1;
        
        #line default
        #line hidden
        
        
        #line 16 "..\..\DataImporter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox From;
        
        #line default
        #line hidden
        
        
        #line 18 "..\..\DataImporter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox To;
        
        #line default
        #line hidden
        
        
        #line 19 "..\..\DataImporter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal MahApps.Metro.Controls.Tile Import;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/TableMaster;component/dataimporter.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\DataImporter.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.Path = ((System.Windows.Controls.TextBox)(target));
            
            #line 12 "..\..\DataImporter.xaml"
            this.Path.TextChanged += new System.Windows.Controls.TextChangedEventHandler(this.Path_TextChanged);
            
            #line default
            #line hidden
            return;
            case 2:
            this.Browser = ((System.Windows.Controls.Button)(target));
            
            #line 13 "..\..\DataImporter.xaml"
            this.Browser.Click += new System.Windows.RoutedEventHandler(this.Browser_Click);
            
            #line default
            #line hidden
            return;
            case 3:
            this.TabControl1 = ((System.Windows.Controls.TabControl)(target));
            return;
            case 4:
            this.From = ((System.Windows.Controls.TextBox)(target));
            
            #line 16 "..\..\DataImporter.xaml"
            this.From.GotFocus += new System.Windows.RoutedEventHandler(this.From_GotFocus);
            
            #line default
            #line hidden
            return;
            case 5:
            this.To = ((System.Windows.Controls.TextBox)(target));
            
            #line 18 "..\..\DataImporter.xaml"
            this.To.GotFocus += new System.Windows.RoutedEventHandler(this.To_GotFocus);
            
            #line default
            #line hidden
            return;
            case 6:
            this.Import = ((MahApps.Metro.Controls.Tile)(target));
            
            #line 19 "..\..\DataImporter.xaml"
            this.Import.Click += new System.Windows.RoutedEventHandler(this.Import_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

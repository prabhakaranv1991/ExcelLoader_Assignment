   M S B A M L   `   `   `  ����L  HSampleApplication, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:SampleApplicationSampleApplication  5   	   A%clr-namespace:SampleApplication.ViewsSampleApplication.Views  5   	   "    SampleApplication.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   1localclr-namespace:SampleApplication     5   	   7views%clr-namespace:SampleApplication.Views     5   	     � Title$  Excel Data Loader��5	   	   $	��500��6#   $	��900��60    � WindowStartupLocation$ CenterScreen=�5
   
    � WindowState$ Normal=�60    � WindowStyle$
 None=�6E    � 
ResizeMode$ NoResize=�6X   
������ .��5      �� -   5        $��XI��5      $��10q�6    J� Effect 5   
   3  -System.Windows.Media.Effects.DropShadowEffect  5         Color$ Black��6      	Direction$	 270Q�5         
BlurRadius$ 10Q�5      	   ShadowDepth$	 3Q�5      ._�5   
   � 
�������.+�5      ]� .+�5      ]� -   5        $��titleBar��5      $	��Top=�5      $��32��5      ���#007daa.+�5      � 
 � RowDefinitions
 5      �� 5      6.    � ColumnDefinitions 5      �� 5      61   �� 5       61   5"      �� 5$       �� Padding$ 8q�6    $��0��6-   $��0��6:   $��Center=�5%       $��COMMODITY DETAILS��5&       
�������$��Light"�5(       $��16(�5)       � 5*      $��0��6   $��1��6(   $��80��68   $��Center=�6C   $��Right=�6^   $��0 0 6 0q�6z   
 5+      �� 5,      6.    5.      �� 5/      61   �� 50      61   .+�53      �� -   5         ��Name$ Minimize��53   !    � ToolTip$ Minimize��63   $��0��6F   $��0��6S   
������ 
�������$��0q�6�   $��
ExtraBlack"�6�   $��-��6�   $��0 0 -2 0q�6�   $��Center=�6�   $��30��6�   6  �� 54      -   5        $ Close��54   !   $ Close��60   $��0��6@   $��1��6M   $��X��6]   
�������
������ $��0q�6�   $��
ExtraBlack"�6�   $��Center=�6�   $��30��6�   6�   55      56      �� 57      -   5        $��content��57   !   58      -   'SampleApplication.Views.ExcelDataLoader  5;      -   5          Name$ ExcelLoaderControl��5;   $   -   'SampleApplication.Views.ExcelDataLoader   Loaded$ ExcelLoaderControl_Load��6@   6c   5>      5@      5A      
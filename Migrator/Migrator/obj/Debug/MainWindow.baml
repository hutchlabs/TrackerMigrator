   M S B A M L   `   `   `  ����T  PMahApps.Metro, Version=1.1.2.0, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  5      fCclr-namespace:MahApps.Metro.Controls.Dialogs;assembly=MahApps.MetroMahApps.Metro.Controls.Dialogs  5       Migrator   Migrator.MainWindow   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      PControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro     5      VDialogCclr-namespace:MahApps.Metro.Controls.Dialogs;assembly=MahApps.Metro     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5        � Title$#  Petra Tracker - Data Migrator��5      (   "MahApps.Metro.Controls.MetroWindow   ShowTitleBar .  � WindowStartupLocation$ CenterScreen=�5
         AccentColorBrush   	GlowBrush# �   5         NonActiveGlowBrush$	 Red��5      $	��400��5   (   $	��290��64    � 
ResizeMode$ NoResize=�5   +   .��5      � $��10,10,10,10q�6    � ColumnDefinitions 5   
   �� 5      $	��120 �6   �� 5      $��* �6    � RowDefinitions 5   
   �� 5      $��60 �6   �� 5      $��60 �6   �� 5      $��60 �6   �� 5      $��60 �6   .+�5   
   �� $��0��6   $��0��6   $��Migration Type��6-   $��Right=�6F   $	��Top=�6b   $��10q�6z   $��26��6�   �� 5   
   -   5         ��Name$ cbx_mt��5   0   $��0��6   $��1��6    $
��Left=�6@   $��
10,17,0,17q�6\   $��Center=�6q   $	��190��6�   $��22��6�   .@�5      �� $��Payments��6   �� 5       $��	Schedules��6   5!      �� 5#   
   $��1��6   $��0��6   $��Last Insert Id��6-   $��Right=�6F   $	��Top=�6b   $��10q�6z   $��26��6�   *   $MahApps.Metro.Controls.NumericUpDown  5$   
   -   5        	  Name$	 lastinsertid��5%   %   $��1��5$   !   $��1��6.   $��Center=�6>   $��	10,0,52,0q�6Y   
   Minimum$
 1Q�5&   %      Interval$ 1Q�61   $	��190��6>   $��26��5'   %   {�.  6Z   �� 5)   
   -   5         ��Name$ btn_migrate��5)      $��2��6&   $��1��63   {�.  $��Migrate��6V   $
��Left=�5*      $��Center=�6-   $	��122��6H   $��36��6T   $��10q�6u   )   #MahApps.Metro.Controls.ProgressRing  5,   
   -   5          Name$ spinner��5,       $��2��61   $��1��6>   $��40��6N   $��48��6Y   $��Right=�6f      IsActive .  $��0,6,55,6q�5-   1    J� RenderTransformOrigin$ 0.55,0.5&�6C   �� 5/   
   -   5         ��Name$ livefeed��5/      $��3��6"   $��0��6/   $��2��6?   $
��Left=�6T   $��Center=�6o   $��26��6�   $	��350��6�   $��10q�6�   50      51      
   M S B A M L   `   `   `  ����   TableMaster    TableMaster.MainWindow   U QMahApps.Metro, Version=1.5.0.23, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4f;Controls,http://metro.mahapps.com/winfx/xaml/controls  R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006        AccentColorBrush(  "MahApps.Metro.Controls.MetroWindow    	GlowBrush#  �    � Title$ Table Master��$	��579��$	��525��$
��Left=�$	��Top=�   RightWindowCommands +  %MahApps.Metro.Controls.WindowCommands  .@��� -    ��Name$
 Help��$
��Help��.��� .+�� .+�!  MahApps.Metro.Controls.Tile     Title$ 总表��$
��Left=�$	��115��$��	10,10,0,0q�$	��Top=�$	��115��$��YouYuan)�$��20(�$
��Bold"�   TitleFontSize$ 22Q����	#CC119EDA
�������   Count$  ��.���� $��79��$��79��$��TableSelect.png��  $
��Left=�$	��115��$��
130,10,0,0q�$	��Top=�$	��115��$ 单科��$ 22Q�$��18(�$��YouYuan)�$
��Bold"�.���� $��90��$��90��$��SingleSubject.png���� -    ��Name$ FileList��$��10,170,96,154q�$��Segoe UI Semibold)�$��16(�
���   �$��1q��� -   $	 Add��$��+��$��30��$��0,135,117,0q�$	��Top=�$��Segoe UI Semibold)�$��22(�$
��Bold"�$��Stretch=�$��0,-4,0,0q�$��Right=�$��45���� -    ��Name$ FileName��$��30��$��10,135,167,0q�	 �� TextWrapping$
	 Wrap=�$	��Top=�$��Microsoft YaHei)�$��18(�$��0,-1,0,0q�&   System.Windows.Media.TextOptions
   TextHintingMode$
 Animated=��� -   $ Select��$��浏览表格��$��30��$��
0,135,10,0q�$	��Top=�$��Microsoft YaHei UI)�$��16(�$��5,4,5,6q�$��Right=�$	��104���� -   $ Remove��$��-��$��0,170,10,154q�$��24(�$��Microsoft YaHei UI)�$
��Bold"�$��5,3,5,6q�$��Right=�$��38��  -     Name$ TemplateSetting��$
��Left=�$	��100��$��	10,0,0,10q�$	��100��$ 编辑模板��$ 18Q�$��Center=�$��YouYuan)�$��18(�$��Bottom=�.���� $��63��$��Settings.png��$
��Fill=�$��64��$	��Top=�$��0,0,0,22q�  -   $ Generate��$��生成��$��115,0,10,10q�$
��Auto��$��YouYuan)�$��48(�$
��Bold"�$	��100��$��Bottom=��� -	   $ Advance��$��名单设置��$��
10,0,0,115q�$��Microsoft YaHei)�$��18(�$��Normal"�$��5,4,5,6q�$��34��$��Bottom=�$
��Left=�$	��100���� -
   $ Up��$	��▲��$��
0,170,53,0q�$��24(�$��Microsoft YaHei UI)�$
��Bold"�$��5,3,5,6q�$��Right=�$��38��$	��112��$	��Top=��� -   $
 Down��$	��▼��$��
0,0,53,154q�$��24(�$��Microsoft YaHei UI)�$
��Bold"�$��5,3,5,6q�$	��112��$��Bottom=�$��Right=�$��38���� $��生成目录：��$��34��$��115,0,0,115q�$��Bottom=�$��18(�$��Microsoft YaHei UI)�$
��Left=�$	��100���� -   $ BroswerPath��$��浏览��$��Right=�$��34��$��
0,0,10,115q�$��Bottom=�$��81��$��16(�$��Microsoft YaHei UI)�$��Normal"��� -   $
 Path��$��34��$��215,0,96,115q�$
	 Wrap=�$��Bottom=�$��Microsoft YaHei)�$��16(�$��0,1,0,0q�+  %MahApps.Metro.Controls.FlyoutsControl  .@�#  MahApps.Metro.Controls.Flyout  -     Name$ 
HelpFlyout��$
��Help��$	��517��$��
0,239,0,-1q�$
��Left=�   Position$ Bottom=����	#FF41B1E1   IsOpen .    -   $ AdvancedFlyout��$��Advanced��$	��450��$��0,0,0,-1q�$
��Left=�$ Right=����	#FF41B1E1 .  .��� $	��471��.+��� $��关键字：��$
��Left=�$��37��$��	10,10,0,0q�$	��Top=�$��78��$��18(�$��Microsoft YaHei UI)��� -   $ KeyWord��$
��Left=�$��31��$��	88,11,0,0q�$
	 Wrap=�$��姓名��$	��Top=�$	��183��$��Microsoft YaHei)�$��16(��� $��	名单：��$
��Left=�$��41��$��	10,61,0,0q�$	��Top=�$	��117��$��18(�$��Microsoft YaHei UI)��� -   $ NameList��$	��230��$��
88,61,46,0q�$	��Top=�$��1q��� -   $ Remove1��$��-��$	��230��$��	0,61,10,0q�$	��Top=�$��Right=�$��36��$��22(�$
��Bold"�$��Microsoft YaHei)�$��5,3,5,6q��� -   $
 Name��$��34��$��88,291,170,0q�$
	 Wrap=�$	��Top=�$��16(�$��Microsoft YaHei)�$��0,2,0,0q��� -   $
 Add1��$��添加��$��Right=�$��34��$��
0,291,90,0q�$	��Top=�$��75��$��Microsoft YaHei UI)�$��14(��� -   $ Import��$��引入��$��Right=�$��34��$��
0,291,10,0q�$	��Top=�$��75��$��Microsoft YaHei UI)�$��14(�
   M S B A M L   `   `   `  ���� J  FBookScrapperDOI, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:BookScrapperDOIBookScrapperDOI       BookScrapperDOI.MainWindow   W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     /localclr-namespace:BookScrapperDOI     X TXceed.Wpf.Toolkit, Version=4.2.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f44xctk)http://schemas.xceed.com/wpf/xaml/toolkit  -     � Title$  
MainWindow��$	��500��$	��760�� � WindowStyle$
 None=� � 
ResizeMode$ NoResize=����#181735.���  J� RenderTransformOrigin$ 0.548,0.496&�$��0,0,10,0q� � RowDefinitions �� $��50 ��� $��50 ��� $��50 ��� $��30 ��� $��30 ��� $��30 ���  � ColumnDefinitions �� $	��760 �.+�� ���#181735$
��Left=�$	��760�� �� �� �� �� .+��� $��DOI Scrapper��
�������$��0��$��Center=�$��Center=�$��Fonts/#Bullying)�$��20(�$��Medium"��� -   $��BtnMinimizeWindow��$��Minimize��$��2��
�������$��14(�$��Medium"�$��157,2,119,-2q�$��2������  �� 
TargetType" ��.���  �� Property � �� �� Value$ Fonts/#Bullying)���  � �� �#C0D514��  �@�Cursor$
 Hand��	 �� Triggers	 P� 
 P� Property
 �@J�IsMouseOver P� Value$
 True��. ���  � �� �#C0C524����  �� 
TargetType" ��.R��� $	��100��$��25�� �� CornerRadius$ 12��#��z��._��� $��Center=�$��Center=��� -   $��BtnCloseWindow��$��Close��$��3��
�������$��14(�$��Medium"�$��
76,2,10,-2q����� " ��.���  � ��$ Fonts/#Bullying)���  � �� �#FE6584��  �@�Cursor$
 Hand��	 P� 
 �@J�IsMouseOver$
 True��. ���  � �� �#FE4C70���� " ��.R��� $	��100��$��25��$ 12��#��z��._��� $��Center=�$��Center=��� $��1��$��3��.+���  �� Orientation$ 
Horizontal=�$��10, 10, 10, 0q�.+��� -   $��BtnSelectExcelFile��$��Select Folder��
�������$��14(�$��Medium"����� " ��.���  � ��$ Fonts/#Bullying)���  � �� �#3DDDC2��  �@�Cursor$
 Hand��	 P� 
 �@J�IsMouseOver$
 True��. ���  � �� �#3DDA50���� " ��.R��� $	��100��$��25��$ 12��#��z��._��� $��Center=�$��Center=��� -   $��TbWorkFolderPath��$��10,0,0,0q����#181735$��1��$��Center=�$��Fonts/#Bullying)�$��14(�$��Medium"�$	��625���� $��2��$��3��.+��� $ 
Horizontal=�$��10, 10, 10, 0q�.+��� -   $��BtnStartWork��$��Start��
�������$��14(�$��Medium"����� " ��.���  � ��$ Fonts/#Bullying)���  � �� �#3DDDC2��  �@�Cursor$
 Hand��	 P� 
 �@J�IsMouseOver$
 True��. ���  � �� �#3DDA50���� " ��.R��� $	��100��$��25��$ 12��#��z��._��� $��Center=�$��Center=��� $��Retry Count��
�������$��Fonts/#Bullying)�$��13(�$��Center=�$��20, 0, 10, 0q�$��Medium"�%  Xceed.Wpf.Toolkit.IntegerUpDown  -     Name$ IudRetryMaxCnt��$��39��$
��14px(��  �Xceed.Wpf.Toolkit.Primitives.UpDownBase`1[[System.Nullable`1[[System.Int32, System.Runtime, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a]], System.Runtime, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a]]   Maximum 20   Minimum 1   Value 3���#3DDDC2
��������� -    ��Name$ CbDomainName��$	��336�����#181735���#3DDDC2$��20,0,0,0q� �� SelectedIndex$ 0��$
��14px(����	#FF3DDDC2.@��� ���#3DDDC2���#181735$��https://sci-hub.ee/���� ���#3DDDC2���#181735$��https://sci-hub.ru/���� ���#3DDDC2���#181735$��https://sci-hub.st/���� ���#3DDDC2���#181735$��https://sci-hub.se/���� ���#3DDDC2���#181735$��https://sci-hub.wf/���� $��3��$ 
Horizontal=�$��0,0,10,0q�.+��� -	    ��Name$ LbProgressStatusTotal��$��0/0 finished��
�������$	��Top=�$
��Left=�$��Fonts/#Bullying)�$��14(�$��Medium"�$��	16,10,0,0q�$��0���� -
   $ LbFailedStatusTotal��$��0 failed��
�������$	��Top=�$
��Left=�$��Fonts/#Bullying)�$��14(�$��Medium"�$��	16,10,0,0q�$��1���� -   $ LbTimeSpeed��$��0 s/file��
�������$	��Top=�$
��Left=�$��Fonts/#Bullying)�$��14(�$��Medium"�$��	16,10,0,0q�$��2���� -   $ LbTimeEstimationTotal��$��0 mins remained��
�������$	��Top=�$
��Left=�$��Fonts/#Bullying)�$��14(�$��Medium"�$��	16,10,0,0q�$��2���� $��4��$ 
Horizontal=�$��0,0,10,0q�.+��� -   $ LbProgressStatus��$��0/0 finished��
�������$	��Top=�$
��Left=�$��Fonts/#Bullying)�$��14(�$��Medium"�$��	16,10,0,0q�$��0���� -   $ LbFailedStatus��$��0 failed��
�������$	��Top=�$
��Left=�$��Fonts/#Bullying)�$��14(�$��Medium"�$��	16,10,0,0q�$��1���� -   $ LbTimeEstimation��$!��No file is under processing��
�������$	��Top=�$
��Left=�$��Fonts/#Bullying)�$��14(�$��Medium"�$��	16,10,0,0q�$��2���� $��5��$ 
Horizontal=�$��	10,0,10,0q�.+�� -    �Name$ PbStatus��$��14��$	��744�����#3DDDC2�� $��6��$ 
Horizontal=�$��	10,0,10,0q�.+��� -    ��Name$ TbLog��$	��246��$�� �� �� TextWrapping$
 Wrap=�
����������#181735$	��733��$��Center=�$��Center=�$
��Auto=�$��Fonts/#Bullying)�$��14(�
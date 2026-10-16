<#
.SYNOPSIS
    Intune Group Exclusion Manager - exclude (or un-exclude) a group on Intune configuration profiles.

.DESCRIPTION
    A four-step GUI:
      1. Connect & load   - choose a tenant, the group and a name filter (e.g. 'CUS' or 'W1x*CUS'),
                            then load. Only matching profiles get their assignments loaded.
      2. Tick profiles    - tick the profiles you want to change.
      3. Choose action    - Add exclusion / Remove exclusion.
                            Rows that will really change turn green.
      4. Review & apply   - preview (WhatIf) first, then apply.
                            Every real run is saved as an audit CSV (before/after).

    Covers the three kinds of configuration profiles:
      - Templates          (deviceManagement/deviceConfigurations)
      - Settings Catalog   (deviceManagement/configurationPolicies)
      - Admin Templates    (deviceManagement/groupPolicyConfigurations)

    Existing assignments (including filters) are kept: the Graph /assign action always
    replaces the full set, so the tool sends the current set plus or minus the exclusion.
    Profiles where the group is *included* are never changed.

    Requires: Windows, PowerShell 5.1 or 7, module Microsoft.Graph.Authentication.

.EXAMPLE
    .\Set-IntuneGroupExclusionGUI.ps1
.EXAMPLE
    .\Set-IntuneGroupExclusionGUI.ps1 -GroupName 'AAD_UG_Update-Ring-00' -NameFilter 'W1x' -TenantId 'contoso.onmicrosoft.com'
#>
[CmdletBinding()]
param(
    [string]$GroupName  = 'AAD_UG_Update-Ring-00',
    [string]$NameFilter = 'W1x',
    [string]$TenantId
)

#region STA check (WPF requires STA)
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe     = (Get-Process -Id $PID).Path
    $argList = @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    if ($GroupName)  { $argList += @('-GroupName', "`"$GroupName`"") }
    if ($NameFilter) { $argList += @('-NameFilter', "`"$NameFilter`"") }
    if ($TenantId)   { $argList += @('-TenantId', $TenantId) }
    Start-Process -FilePath $exe -ArgumentList $argList
    return
}
#endregion

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

#region Data model
if (-not ('ExclGui1.ProfileRow' -as [type])) {
    Add-Type -TypeDefinition @'
using System.ComponentModel;
namespace ExclGui1 {
    public class Obs : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void On(string n) { PropertyChangedEventHandler h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(n)); }
    }
    public class ProfileRow : Obs {
        bool _sel;     public bool     Selected       { get { return _sel; } set { _sel = value; On("Selected"); } }
        bool _can;     public bool     CanApply       { get { return _can; } set { _can = value; On("CanApply"); } }
        string _inc;   public string   Includes       { get { return _inc; } set { _inc = value; On("Includes"); } }
        string _exc;   public string   Excludes       { get { return _exc; } set { _exc = value; On("Excludes"); } }
        string _gs;    public string   GroupState     { get { return _gs; }  set { _gs = value;  On("GroupState"); } }
        string _gst;   public string   GroupStateText { get { return _gst; } set { _gst = value; On("GroupStateText"); } }
        string _st;    public string   Status         { get { return _st; }  set { _st = value;  On("Status"); } }
        string _rs;    public string   RowState       { get { return _rs; }  set { _rs = value;  On("RowState"); } }
        string _tip;   public string   Tip            { get { return _tip; } set { _tip = value; On("Tip"); } }
        public string   Name      { get; set; }
        public string   Type      { get; set; }
        public string   Id        { get; set; }
        public string   Endpoint  { get; set; }
        public bool     Done      { get; set; }
        public bool     PreviewOk { get; set; }
        public string   ErrorText { get; set; }
        public object[] Targets   { get; set; }
    }
    public class AssignItem {
        public string Kind   { get; set; }
        public string Name   { get; set; }
        public string Change { get; set; }
    }
}
'@
}
#endregion

#region XAML
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Intune Group Exclusion Manager" Height="940" Width="1600" MinWidth="1250" MinHeight="760"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="12" Background="#F3F6FA">
  <Window.Resources>
    <BooleanToVisibilityConverter x:Key="b2v"/>

    <Style x:Key="BtnBase" TargetType="Button">
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#1F2937"/>
      <Setter Property="BorderBrush" Value="#C9D2DE"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12,4"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.85"/></Trigger>
              <Trigger Property="IsPressed"   Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.7"/></Trigger>
              <Trigger Property="IsEnabled"   Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Button" BasedOn="{StaticResource BtnBase}"/>
    <Style x:Key="Accent" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#1565C0"/>
      <Setter Property="BorderBrush" Value="#1565C0"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="Go" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#2E7D32"/>
      <Setter Property="BorderBrush" Value="#2E7D32"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="HeaderBtn" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#26FFFFFF"/>
      <Setter Property="BorderBrush" Value="#66FFFFFF"/>
      <Setter Property="Foreground" Value="White"/>
    </Style>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="BorderBrush" Value="#DDE3EA"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12"/>
    </Style>
    <Style x:Key="StepCard" TargetType="Border" BasedOn="{StaticResource Card}">
      <Setter Property="BorderThickness" Value="2"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
    </Style>
    <Style x:Key="Dot" TargetType="Border">
      <Setter Property="Width" Value="30"/>
      <Setter Property="Height" Value="30"/>
      <Setter Property="CornerRadius" Value="15"/>
      <Setter Property="Background" Value="#B0BEC5"/>
      <Setter Property="VerticalAlignment" Value="Top"/>
    </Style>
    <Style x:Key="DotText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="StepTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Foreground" Value="#1F2937"/>
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#6B7280"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Padding="18,12">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
          <GradientStop Color="#0D2F57" Offset="0"/>
          <GradientStop Color="#1565C0" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <DockPanel>
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <Border CornerRadius="12" Background="#26FFFFFF" Padding="10,4" VerticalAlignment="Center">
            <StackPanel Orientation="Horizontal">
              <Ellipse x:Name="connDot" Width="9" Height="9" Fill="#FFB74D" VerticalAlignment="Center"/>
              <TextBlock x:Name="lblConn" Text="Not connected" Foreground="White" Margin="7,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
          </Border>
          <Button x:Name="btnDisconnect" Content="Disconnect" Style="{StaticResource HeaderBtn}" Margin="10,0,0,0" Height="28" IsEnabled="False"/>
        </StackPanel>
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="Intune Group Exclusion Manager" Foreground="White" FontSize="20" FontWeight="SemiBold"/>
          <TextBlock Text="Exclude a group from configuration profiles in four steps" Foreground="#CFE0F5"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <ProgressBar x:Name="progress" Grid.Row="1" Height="4" BorderThickness="0" Background="Transparent" Foreground="#42A5F5" Visibility="Hidden"/>

    <!-- Steps -->
    <UniformGrid Grid.Row="2" Columns="4" Margin="12,8,2,0">

      <!-- Step 1 -->
      <Border x:Name="step1" Style="{StaticResource StepCard}">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
            <Border x:Name="dot1" Style="{StaticResource Dot}"><TextBlock x:Name="num1" Text="1" Style="{StaticResource DotText}"/></Border>
            <StackPanel Margin="10,0,0,0">
              <TextBlock Text="Connect &amp; load" Style="{StaticResource StepTitle}"/>
              <TextBlock x:Name="sub1" Text="Pick the tenant and the group" Style="{StaticResource Hint}"/>
            </StackPanel>
          </DockPanel>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="Tenant" VerticalAlignment="Center" Margin="0,0,8,6"/>
            <ComboBox x:Name="cmbTenant" Grid.Column="1" IsEditable="True" Height="26" Margin="0,0,0,6"
                      ToolTip="Tenant ID or domain (optional). Recently used tenants are remembered."/>
            <TextBlock Grid.Row="1" Text="Group" VerticalAlignment="Center" Margin="0,0,8,6"/>
            <TextBox x:Name="txtGroup" Grid.Row="1" Grid.Column="1" Height="26" VerticalContentAlignment="Center" Margin="0,0,0,6"
                     ToolTip="Display name of the Entra ID group to exclude"/>
            <TextBlock Grid.Row="2" Text="Name" VerticalAlignment="Center" Margin="0,0,8,6"/>
            <Grid Grid.Row="2" Grid.Column="1" Margin="0,0,0,6">
              <TextBox x:Name="txtNameFilter" Height="26" VerticalContentAlignment="Center"
                       ToolTip="Only load profiles whose name contains this text. Use * as wildcard, e.g. CUS or W1x*CUS. Empty = all profiles."/>
              <TextBlock Text="contains... (e.g. CUS or W1x*CUS)" Foreground="#9CA3AF" Margin="7,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"
                         Visibility="{Binding Text.IsEmpty, ElementName=txtNameFilter, Converter={StaticResource b2v}}"/>
            </Grid>
            <Button x:Name="btnLoad" Grid.Row="3" Grid.ColumnSpan="2" Content="Connect &amp; load" Style="{StaticResource Accent}" Height="30"/>
          </Grid>
        </DockPanel>
      </Border>

      <!-- Step 2 -->
      <Border x:Name="step2" Style="{StaticResource StepCard}">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
            <Border x:Name="dot2" Style="{StaticResource Dot}"><TextBlock x:Name="num2" Text="2" Style="{StaticResource DotText}"/></Border>
            <StackPanel Margin="10,0,0,0">
              <TextBlock Text="Tick profiles" Style="{StaticResource StepTitle}"/>
              <TextBlock x:Name="sub2" Text="Tick the profiles you want to change" Style="{StaticResource Hint}"/>
            </StackPanel>
          </DockPanel>
          <StackPanel>
            <TextBlock Text="Use the checkboxes in the list, or quick-select:" Style="{StaticResource Hint}" Margin="0,0,0,8"/>
            <UniformGrid Columns="3">
              <Button x:Name="btnSelNotExcl" Content="Not excluded" Height="30" Margin="0,0,4,0" ToolTip="Tick every visible profile where the group is not assigned yet"/>
              <Button x:Name="btnSelVisible" Content="All visible"  Height="30" Margin="2,0,2,0" ToolTip="Tick every profile that matches the current filters"/>
              <Button x:Name="btnSelNone"    Content="None"         Height="30" Margin="4,0,0,0" ToolTip="Untick everything"/>
            </UniformGrid>
            <TextBlock Text="Tip: Space toggles the highlighted rows." Style="{StaticResource Hint}" Margin="0,8,0,0"/>
          </StackPanel>
        </DockPanel>
      </Border>

      <!-- Step 3 -->
      <Border x:Name="step3" Style="{StaticResource StepCard}">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
            <Border x:Name="dot3" Style="{StaticResource Dot}"><TextBlock x:Name="num3" Text="3" Style="{StaticResource DotText}"/></Border>
            <StackPanel Margin="10,0,0,0">
              <TextBlock Text="Choose action" Style="{StaticResource StepTitle}"/>
              <TextBlock x:Name="sub3" Text="What should happen with the ticked profiles?" Style="{StaticResource Hint}"/>
            </StackPanel>
          </DockPanel>
          <StackPanel>
            <RadioButton x:Name="rbAdd" GroupName="mode" IsChecked="True" Margin="0,0,0,6">
              <StackPanel>
                <TextBlock Text="Add exclusion" FontWeight="SemiBold"/>
                <TextBlock Text="Exclude the group, keep all other assignments" Style="{StaticResource Hint}"/>
              </StackPanel>
            </RadioButton>
            <RadioButton x:Name="rbRemove" GroupName="mode">
              <StackPanel>
                <TextBlock Text="Remove exclusion" FontWeight="SemiBold"/>
                <TextBlock Text="Undo: remove the group exclusion again" Style="{StaticResource Hint}"/>
              </StackPanel>
            </RadioButton>
          </StackPanel>
        </DockPanel>
      </Border>

      <!-- Step 4 -->
      <Border x:Name="step4" Style="{StaticResource StepCard}">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
            <Border x:Name="dot4" Style="{StaticResource Dot}"><TextBlock x:Name="num4" Text="4" Style="{StaticResource DotText}"/></Border>
            <StackPanel Margin="10,0,0,0">
              <TextBlock Text="Review &amp; apply" Style="{StaticResource StepTitle}"/>
              <TextBlock x:Name="sub4" Text="Preview first, then apply" Style="{StaticResource Hint}"/>
            </StackPanel>
          </DockPanel>
          <StackPanel>
            <CheckBox x:Name="chkWhatIf" IsChecked="True" FontWeight="SemiBold">
              <TextBlock Text="Preview only (WhatIf) - change nothing"/>
            </CheckBox>
            <TextBlock Text="Untick to write the changes to the tenant." Style="{StaticResource Hint}" Margin="20,0,0,8"/>
            <Button x:Name="btnApply" Content="Nothing to apply" Style="{StaticResource Accent}" Height="34" IsEnabled="False"/>
            <Button x:Name="btnExport" Content="Export overview" Height="26" Margin="0,6,0,0"/>
          </StackPanel>
        </DockPanel>
      </Border>
    </UniformGrid>

    <!-- Filter bar -->
    <Border Grid.Row="3" Style="{StaticResource Card}" Margin="12,10,12,0" Padding="10,8">
      <DockPanel LastChildFill="False">
        <Grid Width="250">
          <TextBox x:Name="txtFilter" Height="26" VerticalContentAlignment="Center"/>
          <TextBlock Text="Search profiles..." Foreground="#9CA3AF" Margin="7,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"
                     Visibility="{Binding Text.IsEmpty, ElementName=txtFilter, Converter={StaticResource b2v}}"/>
        </Grid>
        <TextBlock Text="Type" VerticalAlignment="Center" Margin="14,0,6,0"/>
        <ComboBox x:Name="cmbType" Width="160" Height="26"/>
        <TextBlock Text="Group" VerticalAlignment="Center" Margin="14,0,6,0"/>
        <ComboBox x:Name="cmbState" Width="140" Height="26"/>
        <CheckBox x:Name="chkTicked"  Content="Ticked only"      VerticalAlignment="Center" Margin="16,0,0,0"/>
        <CheckBox x:Name="chkChanges" Content="Will change only" VerticalAlignment="Center" Margin="14,0,0,0"/>
        <TextBlock x:Name="lblCount" DockPanel.Dock="Right" VerticalAlignment="Center" Foreground="#4B5563"/>
      </DockPanel>
    </Border>

    <!-- Main area -->
    <Grid Grid.Row="4" Margin="12,10,12,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="10"/>
        <ColumnDefinition Width="340"/>
      </Grid.ColumnDefinitions>

      <Border Style="{StaticResource Card}" Padding="0">
        <DataGrid x:Name="grid" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False"
                  HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#EEF1F5"
                  BorderThickness="0" Background="White" RowBackground="White" MinRowHeight="28"
                  SelectionMode="Extended" SelectionUnit="FullRow" EnableRowVirtualization="True">
          <DataGrid.ContextMenu>
            <ContextMenu>
              <MenuItem x:Name="mnuTick"   Header="Tick highlighted rows"/>
              <MenuItem x:Name="mnuUntick" Header="Untick highlighted rows"/>
              <Separator/>
              <MenuItem x:Name="mnuCopy"   Header="Copy profile name(s)"/>
            </ContextMenu>
          </DataGrid.ContextMenu>
          <DataGrid.ColumnHeaderStyle>
            <Style TargetType="DataGridColumnHeader">
              <Setter Property="Background" Value="#F8FAFC"/>
              <Setter Property="FontWeight" Value="SemiBold"/>
              <Setter Property="Foreground" Value="#374151"/>
              <Setter Property="Padding" Value="6,7"/>
              <Setter Property="BorderBrush" Value="#E5E7EB"/>
              <Setter Property="BorderThickness" Value="0,0,0,1"/>
            </Style>
          </DataGrid.ColumnHeaderStyle>
          <DataGrid.CellStyle>
            <Style TargetType="DataGridCell">
              <Setter Property="BorderThickness" Value="0"/>
              <Setter Property="VerticalContentAlignment" Value="Center"/>
              <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                  <Setter Property="Background" Value="Transparent"/>
                  <Setter Property="Foreground" Value="#111827"/>
                  <Setter Property="BorderBrush" Value="Transparent"/>
                </Trigger>
              </Style.Triggers>
            </Style>
          </DataGrid.CellStyle>
          <DataGrid.RowStyle>
            <Style TargetType="DataGridRow">
              <Setter Property="ToolTip" Value="{Binding Tip}"/>
              <Setter Property="BorderThickness" Value="4,0,0,0"/>
              <Setter Property="BorderBrush" Value="Transparent"/>
              <Style.Triggers>
                <DataTrigger Binding="{Binding RowState}" Value="Change">
                  <Setter Property="Background" Value="#E6F6EA"/>
                  <Setter Property="BorderBrush" Value="#2E7D32"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding RowState}" Value="Done">
                  <Setter Property="Background" Value="#C9EED2"/>
                  <Setter Property="BorderBrush" Value="#1B5E20"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding RowState}" Value="NoChange">
                  <Setter Property="Background" Value="#F5F7FA"/>
                  <Setter Property="BorderBrush" Value="#B0BEC5"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding RowState}" Value="Error">
                  <Setter Property="Background" Value="#FDE7E7"/>
                  <Setter Property="BorderBrush" Value="#C62828"/>
                </DataTrigger>
                <MultiDataTrigger>
                  <MultiDataTrigger.Conditions>
                    <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsSelected}" Value="True"/>
                    <Condition Binding="{Binding RowState}" Value="Idle"/>
                  </MultiDataTrigger.Conditions>
                  <Setter Property="Background" Value="#E8F1FC"/>
                </MultiDataTrigger>
                <Trigger Property="IsSelected" Value="True">
                  <Setter Property="BorderBrush" Value="#1565C0"/>
                </Trigger>
              </Style.Triggers>
            </Style>
          </DataGrid.RowStyle>
          <DataGrid.Columns>
            <DataGridTemplateColumn Width="42" SortMemberPath="Selected">
              <DataGridTemplateColumn.Header>
                <CheckBox x:Name="chkHeaderAll" ToolTip="Tick / untick all visible rows"/>
              </DataGridTemplateColumn.Header>
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                            IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTextColumn Header="Configuration profile" Binding="{Binding Name}" Width="3*">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock"><Setter Property="TextTrimming" Value="CharacterEllipsis"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="120">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock"><Setter Property="Foreground" Value="#6B7280"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="Included" Binding="{Binding Includes}" Width="2*">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="Excluded" Binding="{Binding Excludes}" Width="2*">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTemplateColumn Header="Group" Width="120" SortMemberPath="GroupState">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse x:Name="gdot" Width="8" Height="8" Fill="#B0BEC5" VerticalAlignment="Center"/>
                    <TextBlock Text="{Binding GroupStateText}" Margin="6,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <DataTemplate.Triggers>
                    <DataTrigger Binding="{Binding GroupState}" Value="Excluded"><Setter TargetName="gdot" Property="Fill" Value="#22C55E"/></DataTrigger>
                    <DataTrigger Binding="{Binding GroupState}" Value="Included"><Setter TargetName="gdot" Property="Fill" Value="#EF4444"/></DataTrigger>
                  </DataTemplate.Triggers>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="1.4*">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
          </DataGrid.Columns>
        </DataGrid>
      </Border>

      <!-- Assignments panel -->
      <Border Grid.Column="2" Style="{StaticResource Card}">
        <DockPanel>
          <TextBlock DockPanel.Dock="Top" Text="Assignments" Style="{StaticResource StepTitle}"/>
          <TextBlock x:Name="lblSel" DockPanel.Dock="Top" Text="Highlight a row" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,6,0,0"/>
          <TextBlock x:Name="lblSelInfo" DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,2,0,8"/>
          <TextBlock DockPanel.Dock="Bottom" Style="{StaticResource Hint}" Margin="0,8,0,0"
                     Text="Green = will be added, red = will be removed. Filters on existing assignments are kept."/>
          <ListBox x:Name="lstAssign" HorizontalContentAlignment="Stretch" BorderBrush="#E5E7EB">
            <ListBox.ItemTemplate>
              <DataTemplate>
                <DockPanel Margin="2,3">
                  <Border x:Name="kindBd" DockPanel.Dock="Left" CornerRadius="9" Padding="7,1" Background="#E5E7EB" VerticalAlignment="Center">
                    <TextBlock x:Name="kindTx" Text="{Binding Kind}" FontSize="11" Foreground="#374151"/>
                  </Border>
                  <TextBlock x:Name="nm" Text="{Binding Name}" Margin="8,0,0,0" TextWrapping="Wrap" VerticalAlignment="Center"/>
                </DockPanel>
                <DataTemplate.Triggers>
                  <DataTrigger Binding="{Binding Kind}" Value="Include">
                    <Setter TargetName="kindBd" Property="Background" Value="#DCEBFB"/>
                    <Setter TargetName="kindTx" Property="Foreground" Value="#0D47A1"/>
                  </DataTrigger>
                  <DataTrigger Binding="{Binding Kind}" Value="Exclude">
                    <Setter TargetName="kindBd" Property="Background" Value="#FFE9CC"/>
                    <Setter TargetName="kindTx" Property="Foreground" Value="#8A4B00"/>
                  </DataTrigger>
                  <DataTrigger Binding="{Binding Change}" Value="Add">
                    <Setter TargetName="nm" Property="Foreground" Value="#1B5E20"/>
                    <Setter TargetName="nm" Property="FontWeight" Value="SemiBold"/>
                  </DataTrigger>
                  <DataTrigger Binding="{Binding Change}" Value="Remove">
                    <Setter TargetName="nm" Property="Foreground" Value="#B71C1C"/>
                    <Setter TargetName="nm" Property="TextDecorations" Value="Strikethrough"/>
                  </DataTrigger>
                </DataTemplate.Triggers>
              </DataTemplate>
            </ListBox.ItemTemplate>
          </ListBox>
        </DockPanel>
      </Border>
    </Grid>

    <!-- Log -->
    <Expander Grid.Row="5" Header="Activity log" IsExpanded="True" Margin="12,8,12,10">
      <TextBox x:Name="txtLog" Height="110" IsReadOnly="True" FontFamily="Consolas" FontSize="11" Margin="0,4,0,0"
               VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Background="White" BorderBrush="#DDE3EA"/>
    </Expander>
  </Grid>
</Window>
'@
$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$ui = @{}
foreach ($n in 'connDot','lblConn','btnDisconnect','progress',
               'step1','dot1','num1','sub1','cmbTenant','txtGroup','txtNameFilter','btnLoad',
               'step2','dot2','num2','sub2','btnSelNotExcl','btnSelVisible','btnSelNone',
               'step3','dot3','num3','sub3','rbAdd','rbRemove',
               'step4','dot4','num4','sub4','chkWhatIf','btnApply','btnExport',
               'txtFilter','cmbType','cmbState','chkTicked','chkChanges','lblCount',
               'grid','lblSel','lblSelInfo','lstAssign','txtLog',
               'mnuTick','mnuUntick','mnuCopy','chkHeaderAll') {
    $ui[$n] = $window.FindName($n)
}
# Fallbacks: header and context menu content is not always in the window namescope
if (-not $ui.chkHeaderAll -and $ui.grid.Columns[0].Header -is [System.Windows.Controls.CheckBox]) {
    $ui.chkHeaderAll = $ui.grid.Columns[0].Header
}
if (-not $ui.mnuTick -and $ui.grid.ContextMenu) {
    $mi = @($ui.grid.ContextMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] })
    $ui.mnuTick = $mi[0]; $ui.mnuUntick = $mi[1]; $ui.mnuCopy = $mi[2]
}
#endregion

#region State & helpers
$graph                  = 'https://graph.microsoft.com/beta'
$graphV1                = 'https://graph.microsoft.com/v1.0'
$script:Rows            = New-Object 'System.Collections.ObjectModel.ObservableCollection[ExclGui1.ProfileRow]'
$script:AssignItems     = New-Object 'System.Collections.ObjectModel.ObservableCollection[ExclGui1.AssignItem]'
$script:GroupNames      = @{}   # group id -> display name
$script:GroupId         = ''
$script:GroupName       = ''
$script:GroupDirty      = $false
$script:LogRows         = New-Object System.Collections.Generic.List[object]
$script:Busy            = $false
$script:Connected       = $false
$script:TenantLabel     = ''
$script:TenantInput     = ''
$script:Mode            = 'Add'
$script:LastApplyOk     = $false
$script:WorkerRs        = $null    # background runspace for all Graph calls (keeps the window responsive)
$script:CancelRequested = $false
$script:Cancellable     = $false
$script:SettingsPath    = Join-Path $env:LOCALAPPDATA 'IntuneGroupExclusionManager\settings.json'
$script:BrushConv       = New-Object System.Windows.Media.BrushConverter

$ui.grid.ItemsSource      = $script:Rows
$ui.lstAssign.ItemsSource = $script:AssignItems
$script:View = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:Rows)

$collections = @(
    @{ Type = 'Template';         Uri = 'deviceManagement/deviceConfigurations';      Name = 'displayName' }
    @{ Type = 'Settings Catalog'; Uri = 'deviceManagement/configurationPolicies';     Name = 'name' }
    @{ Type = 'Admin Template';   Uri = 'deviceManagement/groupPolicyConfigurations'; Name = 'displayName' }
)

function Update-UI { $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background) }

function Get-Brush { param([string]$Hex) $script:BrushConv.ConvertFromString($Hex) }

function Write-Log {
    param([string]$Text, [string]$Level = 'INFO')
    $ui.txtLog.AppendText(('{0} [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Text) + [Environment]::NewLine)
    $ui.txtLog.ScrollToEnd()
    Update-UI
}

function Get-GraphError {
    param($Err)
    if ($Err.ErrorDetails -and $Err.ErrorDetails.Message) {
        try { return ($Err.ErrorDetails.Message | ConvertFrom-Json).error.message } catch { return $Err.ErrorDetails.Message }
    }
    return $Err.Exception.Message
}

#   All Microsoft Graph work (module import, sign-in, requests) runs in a background
#   runspace. The UI thread only waits and keeps pumping window messages, so the
#   window never freezes - also not while the sign-in dialog is open.
function Wait-Pipeline {
    param($Ps)
    $h = $Ps.BeginInvoke()
    while (-not $h.AsyncWaitHandle.WaitOne(25)) {
        Update-UI
        if ($script:CancelRequested) {
            $script:CancelRequested = $false
            try { [void]$Ps.BeginStop($null, $null) } catch {}
            $script:WorkerRs = $null     # abandon the busy runspace; a fresh one is created next time
            throw 'Cancelled by user.'
        }
    }
    try { $out = $Ps.EndInvoke($h) }
    catch {
        $inner = $_.Exception.InnerException
        if ($inner -is [System.Management.Automation.RuntimeException] -and $inner.ErrorRecord) { throw $inner.ErrorRecord }
        if ($inner) { throw $inner }
        throw
    }
    if ($Ps.Streams.Error.Count -gt 0) { throw $Ps.Streams.Error[0] }
    return ,([object[]]@($out))
}

function Get-Worker {
    if ($script:WorkerRs -and $script:WorkerRs.RunspaceStateInfo.State -eq 'Opened') { return $script:WorkerRs }
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript('$ErrorActionPreference = "Stop"; Import-Module Microsoft.Graph.Authentication -ErrorAction Stop')
    try { [void](Wait-Pipeline $ps) }
    finally { if ($ps.InvocationStateInfo.State -notin 'Running','Stopping') { $ps.Dispose() } }
    $script:WorkerRs = $rs
    return $rs
}

function Invoke-Worker {
    param([scriptblock]$Script, [hashtable]$Params = @{})
    $ps = [PowerShell]::Create()
    $ps.Runspace = Get-Worker
    [void]$ps.AddScript($Script.ToString())
    foreach ($k in $Params.Keys) { [void]$ps.AddParameter($k, $Params[$k]) }
    try { $res = Wait-Pipeline $ps }
    finally { if ($ps.InvocationStateInfo.State -notin 'Running','Stopping') { $ps.Dispose() } }
    return ,$res
}

function Invoke-Graph {
    param([string]$Method = 'GET', [string]$Uri, [string]$Body)
    $res = Invoke-Worker -Params @{ Method = $Method; Uri = $Uri; Body = $Body } -Script {
        param($Method, $Uri, $Body)
        if ($Body) { Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType 'application/json' -OutputType PSObject }
        else       { Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject }
    }
    if ($res.Count -gt 0) { return $res[0] }
}

function Get-GraphAll {
    param([string]$Uri)
    $res = Invoke-Worker -Params @{ Uri = $Uri } -Script {
        param($Uri)
        $next = $Uri
        while ($next) {
            $r = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
            if ($r.value) { $r.value }
            $next = $r.'@odata.nextLink'
        }
    }
    # Emit the items one by one, so @(Get-GraphAll ...) is a flat list of items
    foreach ($i in $res) { $i }
}

function Get-GraphContext {
    $res = Invoke-Worker -Script {
        $c = Get-MgContext
        if ($c) { [pscustomobject]@{ TenantId = [string]$c.TenantId; Account = [string]$c.Account } }
    }
    if ($res.Count -gt 0) { return $res[0] }
}

#   Sign-in: interactive authorization code flow with PKCE (no device code). The tool opens the
#   normal Microsoft sign-in page in the default browser (MFA / Conditional Access apply as usual)
#   and receives the answer on a temporary http://localhost port. The token is handed to the Graph
#   module with Connect-MgGraph -AccessToken. This avoids the WAM sign-in broker, which cannot show
#   its window from the background runspace.
$script:ClientId   = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools
$script:AuthScopes = 'https://graph.microsoft.com/DeviceManagementConfiguration.ReadWrite.All https://graph.microsoft.com/Group.Read.All offline_access openid profile'
$script:Token      = $null
$script:TokenInfo  = $null
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function ConvertTo-Base64Url { param([byte[]]$Bytes) [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }

function Get-JwtClaims {
    param([string]$Jwt)
    $p = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
    switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
}

function Get-OAuthError {
    param($Err)
    try { $j = $Err.ErrorDetails.Message | ConvertFrom-Json; if ($j.error_description) { return ($j.error_description -split "`r?`n")[0] } } catch {}
    return $Err.Exception.Message
}

function Get-InteractiveToken {
    param([string]$Tenant)
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    $buf = New-Object byte[] 32; $rng.GetBytes($buf); $verifier = ConvertTo-Base64Url $buf
    $rng.GetBytes($buf); $state = ConvertTo-Base64Url $buf
    $sha = [Security.Cryptography.SHA256]::Create()
    $challenge = ConvertTo-Base64Url ($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier)))

    # Free local port for the redirect
    $tcp = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0); $tcp.Start()
    $port = $tcp.LocalEndpoint.Port; $tcp.Stop()
    $redirect = "http://localhost:$port"
    $listener = New-Object Net.HttpListener
    $listener.Prefixes.Add("$redirect/")
    $listener.Start()
    try {
        $q = @{
            client_id = $script:ClientId; response_type = 'code'; redirect_uri = $redirect; response_mode = 'query'
            scope = $script:AuthScopes; state = $state; code_challenge = $challenge; code_challenge_method = 'S256'; prompt = 'select_account'
        }
        $url = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/authorize?" + (($q.Keys | ForEach-Object { "$_=$([uri]::EscapeDataString([string]$q[$_]))" }) -join '&')
        Write-Log 'Opening the Microsoft sign-in page in your default browser... Complete the sign-in there and come back to this window.'
        Start-Process $url

        $task = $listener.GetContextAsync()
        while (-not $task.Wait(100)) {
            Update-UI
            if ($script:CancelRequested) { $script:CancelRequested = $false; throw 'Cancelled by user.' }
        }
        $ctx = $task.Result
        $code = $ctx.Request.QueryString['code']; $err = $ctx.Request.QueryString['error']
        $okState = $ctx.Request.QueryString['state'] -eq $state
        $msg = if ($code -and $okState) { 'Signed in. You can close this tab and return to Intune Group Exclusion Manager.' }
               else { 'Sign-in failed: ' + [Net.WebUtility]::HtmlEncode([string]$ctx.Request.QueryString['error_description']) }
        $html = [Text.Encoding]::UTF8.GetBytes("<html><body style='font-family:Segoe UI;padding:40px'><h3>$msg</h3></body></html>")
        $ctx.Response.ContentType = 'text/html; charset=utf-8'
        $ctx.Response.OutputStream.Write($html, 0, $html.Length); $ctx.Response.Close()
        if ($err)      { throw "Sign-in failed: $($ctx.Request.QueryString['error_description'])" }
        if (-not $okState -or -not $code) { throw 'Sign-in failed: invalid response from the sign-in page.' }
    }
    finally { try { $listener.Stop(); $listener.Close() } catch {} }

    try {
        return Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body @{
            client_id = $script:ClientId; grant_type = 'authorization_code'; code = $code; redirect_uri = $redirect
            code_verifier = $verifier; scope = $script:AuthScopes
        }
    } catch { throw "Token request failed: $(Get-OAuthError $_)" }
}

function Set-GraphToken {
    # Hand a token to the Graph module in the worker runspace and remember who signed in
    param($Tok)
    $claims = Get-JwtClaims $Tok.access_token
    $script:Token = [pscustomobject]@{
        Access  = [string]$Tok.access_token
        Refresh = [string]$Tok.refresh_token
        Expires = (Get-Date).AddSeconds([int]$Tok.expires_in - 60)
        Tenant  = [string]$claims.tid
    }
    $account = if ($claims.upn) { $claims.upn } elseif ($claims.unique_name) { $claims.unique_name } else { $claims.name }
    $script:TokenInfo = [pscustomobject]@{ Account = [string]$account; TenantId = [string]$claims.tid }
    $missing = @('DeviceManagementConfiguration.ReadWrite.All', 'Group.Read.All' | Where-Object { @(([string]$claims.scp) -split ' ') -notcontains $_ })
    if ($missing) { Write-Log "The token is missing the permission(s): $($missing -join ', '). Admin consent may be required." 'WARN' }
    [void](Invoke-Worker -Params @{ T = $script:Token.Access } -Script {
        param($T)
        Connect-MgGraph -AccessToken (ConvertTo-SecureString $T -AsPlainText -Force) -NoWelcome | Out-Null
    })
}

function Update-TokenIfNeeded {
    # Silently renew the access token (about 1 hour valid) with the refresh token
    if (-not $script:Token -or (Get-Date) -lt $script:Token.Expires.AddMinutes(-5)) { return }
    if (-not $script:Token.Refresh) { throw 'The sign-in has expired. Click Disconnect and sign in again.' }
    Write-Log 'Renewing the access token...'
    try {
        $tok = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($script:Token.Tenant)/oauth2/v2.0/token" -Body @{
            client_id = $script:ClientId; grant_type = 'refresh_token'; refresh_token = $script:Token.Refresh; scope = $script:AuthScopes
        }
    } catch { throw "The sign-in could not be renewed ($(Get-OAuthError $_)). Click Disconnect and sign in again." }
    Set-GraphToken $tok
}

function Get-HighlightedRows {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($i in $ui.grid.SelectedItems) { if ($i -is [ExclGui1.ProfileRow]) { $list.Add($i) } }
    return ,$list.ToArray()
}

function Get-VisibleRows {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($i in $script:View) { $list.Add($i) }
    return ,$list.ToArray()
}
#endregion

#region Assignment helpers
function Test-Exclusion { param($Target) ([string]$Target.'@odata.type') -like '*exclusionGroupAssignmentTarget' }

function Get-TargetName {
    param($Target)
    $type = [string]$Target.'@odata.type'
    $name = switch -Wildcard ($type) {
        '*allDevicesAssignmentTarget'       { 'All devices' }
        '*allLicensedUsersAssignmentTarget' { 'All users' }
        default {
            $id = [string]$Target.groupId
            if ($script:GroupNames.ContainsKey($id)) { $script:GroupNames[$id] } elseif ($id) { "(unknown group) $id" } else { $type }
        }
    }
    $ft = [string]$Target.deviceAndAppManagementAssignmentFilterType
    if ($ft -and $ft -ne 'none') { $name += " [filter: $ft]" }
    return $name
}

function Get-GroupState {
    param($Row)
    if (-not $script:GroupId) { return 'None' }
    foreach ($t in @($Row.Targets)) {
        if ([string]$t.groupId -ne $script:GroupId) { continue }
        if (Test-Exclusion $t) { return 'Excluded' } else { return 'Included' }
    }
    return 'None'
}

function Get-PlannedTargets {
    # The full set of targets after the step 3 action
    param($Row)
    $cur = @($Row.Targets | Where-Object { $null -ne $_ })
    if ($script:Mode -eq 'Add') {
        return @($cur + [pscustomobject]@{ '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'; groupId = $script:GroupId })
    }
    return @($cur | Where-Object { -not ((Test-Exclusion $_) -and [string]$_.groupId -eq $script:GroupId) })
}

function Format-Targets {
    param($Targets, [bool]$Exclusions)
    $names = @($Targets | Where-Object { $null -ne $_ -and (Test-Exclusion $_) -eq $Exclusions } | ForEach-Object { Get-TargetName $_ })
    return (($names | Sort-Object) -join ', ')
}

function Resolve-GroupNames {
    # Group IDs -> display names (15 per request, the limit for 'id in (...)')
    param([string[]]$Ids)
    $todo = @($Ids | Where-Object { $_ -and -not $script:GroupNames.ContainsKey($_) } | Sort-Object -Unique)
    for ($i = 0; $i -lt $todo.Count; $i += 15) {
        $chunk  = $todo[$i..([Math]::Min($i + 14, $todo.Count - 1))]
        $filter = [uri]::EscapeDataString("id in ('" + ($chunk -join "','") + "')")
        foreach ($g in (Get-GraphAll "$graphV1/groups?`$filter=$filter&`$select=id,displayName")) {
            $script:GroupNames[[string]$g.id] = [string]$g.displayName
        }
    }
}

function Get-AssignmentsBatch {
    # Profile id -> assignment targets, via Graph JSON batching (20 requests per call, throttled ones retried)
    param($Profiles)
    $map  = @{}
    $todo = [object[]]@($Profiles | ForEach-Object { $_ })   # @($List) fails in PS 5.1 ("Argument types do not match")
    Start-Progress $todo.Count
    for ($attempt = 1; $todo.Count -and $attempt -le 4; $attempt++) {
        $retry = New-Object System.Collections.Generic.List[object]
        $wait  = 0
        for ($i = 0; $i -lt $todo.Count; $i += 20) {
            $chunk = @($todo[$i..([Math]::Min($i + 19, $todo.Count - 1))])
            $reqs  = for ($j = 0; $j -lt $chunk.Count; $j++) {
                @{ id = "$j"; method = 'GET'; url = "/$($chunk[$j].Endpoint)/$($chunk[$j].Id)/assignments" }
            }
            $resp = Invoke-Graph -Method POST -Uri "$graph/`$batch" -Body (@{ requests = @($reqs) } | ConvertTo-Json -Depth 5)
            foreach ($r in @($resp.responses)) {
                $p = $chunk[[int]$r.id]
                if ([int]$r.status -eq 200) {
                    $map[$p.Id] = [object[]]@($r.body.value | Where-Object { $_.target } | ForEach-Object { $_.target })
                    Step-Progress
                } elseif ([int]$r.status -eq 429 -or [int]$r.status -ge 500) {
                    $retry.Add($p)
                    $ra = 0; if ($r.headers -and $r.headers.'Retry-After') { [int]::TryParse([string]$r.headers.'Retry-After', [ref]$ra) | Out-Null }
                    $wait = [Math]::Max($wait, [Math]::Max($ra, 2))
                } else {
                    Write-Log "Assignments of '$($p.Name)' could not be read: $($r.body.error.message)" 'WARN'
                    Step-Progress
                }
            }
        }
        $todo = $retry.ToArray()
        if ($todo.Count) {
            Write-Log "Throttled by Graph - retrying $($todo.Count) profiles in $wait s..." 'WARN'
            $until = (Get-Date).AddSeconds($wait)
            while ((Get-Date) -lt $until) { Start-Sleep -Milliseconds 100; Update-UI; if ($script:CancelRequested) { throw 'Cancelled by user.' } }
        }
    }
    foreach ($p in $todo) { Write-Log "Assignments of '$($p.Name)' could not be read (throttled)." 'WARN' }
    Stop-Progress
    return $map
}
#endregion

#region Settings & audit log
function Read-Settings {
    try {
        if (Test-Path $script:SettingsPath) { return (Get-Content -Path $script:SettingsPath -Raw | ConvertFrom-Json) }
    } catch {}
    return [pscustomobject]@{ LastGroup = ''; LastNameFilter = $null; RecentTenants = @() }
}

function Update-TenantCombo {
    $txt = $ui.cmbTenant.Text
    $ui.cmbTenant.Items.Clear()
    foreach ($t in @((Read-Settings).RecentTenants)) { if ($t) { [void]$ui.cmbTenant.Items.Add([string]$t) } }
    $ui.cmbTenant.Text = $txt
}

function Save-Settings {
    param([string]$Group, [string]$NameFilter, [string]$Tenant)
    try {
        $s = Read-Settings
        $recent = @(@($s.RecentTenants) | Where-Object { $_ -and $_ -ne $Tenant })
        if ($Tenant) { $recent = @($Tenant) + $recent }
        $obj = [pscustomobject]@{ LastGroup = $Group; LastNameFilter = $NameFilter; RecentTenants = @($recent | Select-Object -First 10) }
        $dir = Split-Path $script:SettingsPath
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $obj | ConvertTo-Json | Set-Content -Path $script:SettingsPath -Encoding UTF8
        Update-TenantCombo
    } catch {}
}

function Get-LogDir {
    $candidates = @()
    if ($PSScriptRoot) { $candidates += (Join-Path $PSScriptRoot 'Logs') }
    $candidates += (Join-Path $env:LOCALAPPDATA 'IntuneGroupExclusionManager\Logs')
    foreach ($d in $candidates) {
        try {
            if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null }
            $probe = Join-Path $d '.write-test'
            Set-Content -Path $probe -Value 'x' -ErrorAction Stop
            Remove-Item -Path $probe -ErrorAction SilentlyContinue
            return $d
        } catch {}
    }
    return $env:TEMP
}
#endregion

#region Steps, progress & busy state
function Set-StepState {
    param([int]$N, [string]$State, [string]$Sub)
    $card = $ui["step$N"]; $dot = $ui["dot$N"]; $num = $ui["num$N"]
    switch ($State) {
        'done' {
            $dot.Background = Get-Brush '#2E7D32'; $num.Text = [string][char]0x2713
            $card.BorderBrush = Get-Brush '#A5D6A7'; $card.Background = Get-Brush '#FFFFFF'
        }
        'active' {
            $dot.Background = Get-Brush '#1565C0'; $num.Text = "$N"
            $card.BorderBrush = Get-Brush '#1565C0'; $card.Background = Get-Brush '#F5F9FF'
        }
        default {
            $dot.Background = Get-Brush '#B0BEC5'; $num.Text = "$N"
            $card.BorderBrush = Get-Brush '#DDE3EA'; $card.Background = Get-Brush '#FFFFFF'
        }
    }
    if ($Sub) { $ui["sub$N"].Text = $Sub }
}

function Update-Steps {
    param([int]$Selected, [int]$Changes)
    $loaded = $script:Rows.Count -gt 0
    if ($loaded -and $script:GroupDirty) {
        Set-StepState 1 'active' 'Group changed - click Connect & load again'
    } elseif ($loaded) {
        $f = if ($script:LoadFilter) { " matching '$($script:LoadFilter)'" } else { '' }
        Set-StepState 1 'done' "$($script:Rows.Count) profiles$f - group: $($script:GroupName)"
    } else {
        Set-StepState 1 'active' 'Pick the tenant and the group'
    }

    if (-not $loaded)        { Set-StepState 2 'todo'   'Tick the profiles you want to change' }
    elseif ($Selected -gt 0) { Set-StepState 2 'done'   "$Selected profiles ticked" }
    else                     { Set-StepState 2 'active' 'Tick the profiles you want to change' }

    if (-not $loaded -or $Selected -eq 0) { Set-StepState 3 'todo'   'What should happen with the ticked profiles?' }
    elseif ($Changes -gt 0)               { Set-StepState 3 'done'   "$Changes rows will change (green)" }
    else                                  { Set-StepState 3 'active' 'The ticked profiles need no change with this action' }

    if ($Changes -gt 0)                       { Set-StepState 4 'active' 'Preview first, then apply' }
    elseif ($script:LastApplyOk -and $loaded) { Set-StepState 4 'done'   'Changes applied - audit log saved' }
    else                                      { Set-StepState 4 'todo'   'Preview first, then apply' }
}

function Update-Summary {
    $sel = 0; $chg = 0; $excl = 0
    foreach ($r in $script:Rows) {
        if ($r.Selected) { $sel++; if ($r.CanApply) { $chg++ } }
        if ($r.GroupState -eq 'Excluded') { $excl++ }
    }
    $visible = 0; foreach ($r in $script:View) { $visible++ }
    $ui.lblCount.Text = "$visible of $($script:Rows.Count) visible   |   $excl excluded   |   $sel ticked   |   $chg will change"

    $whatIf = [bool]$ui.chkWhatIf.IsChecked
    $ui.btnApply.Content   = if ($chg -eq 0) { 'Nothing to apply' } elseif ($whatIf) { "Preview $chg changes" } else { "Apply $chg changes" }
    $ui.btnApply.Style     = if ($whatIf) { $window.FindResource('Accent') } else { $window.FindResource('Go') }
    $ui.btnApply.IsEnabled = ($chg -gt 0) -and -not $script:Busy -and $script:Connected -and -not $script:GroupDirty
    Update-Steps -Selected $sel -Changes $chg
}

function Set-Busy {
    param([bool]$On)
    $script:Busy = $On
    foreach ($b in 'btnLoad','btnExport','btnSelNotExcl','btnSelVisible','btnSelNone') { $ui[$b].IsEnabled = -not $On }
    foreach ($r in 'rbAdd','rbRemove','chkWhatIf','grid','txtGroup','txtNameFilter','cmbTenant') { $ui[$r].IsEnabled = -not $On }
    $ui.btnDisconnect.IsEnabled = (-not $On) -and $script:Connected
    if (-not $On) {
        $script:Cancellable = $false; $script:CancelRequested = $false
        $ui.btnLoad.Content = 'Connect & load'
    }
    $window.Cursor = if ($On) { [System.Windows.Input.Cursors]::Wait } else { $null }
    Update-Summary
    Update-UI
}

function Start-Progress { param([int]$Max) $ui.progress.Minimum = 0; $ui.progress.Maximum = [Math]::Max(1, $Max); $ui.progress.Value = 0; $ui.progress.Visibility = 'Visible'; Update-UI }
function Step-Progress  { $ui.progress.Value = $ui.progress.Value + 1; Update-UI }
function Stop-Progress  { $ui.progress.Visibility = 'Hidden' }
#endregion

#region Row and panel logic
function Update-Row {
    param($Row, [switch]$KeepFlags)
    if (-not $KeepFlags) { $Row.ErrorText = ''; $Row.PreviewOk = $false }

    $state = Get-GroupState $Row
    $Row.GroupState     = $state
    $Row.GroupStateText = switch ($state) { 'Excluded' { 'Excluded' } 'Included' { 'Included' } default { 'Not assigned' } }
    $Row.Includes       = Format-Targets $Row.Targets $false
    $Row.Excludes       = Format-Targets $Row.Targets $true

    $Row.CanApply = [bool]$script:GroupId -and $(if ($script:Mode -eq 'Add') { $state -eq 'None' } else { $state -eq 'Excluded' })
    if ($Row.CanApply) { $Row.Done = $false }
    $Row.Tip = if ($Row.CanApply) { 'After apply - excluded: ' + (Format-Targets (Get-PlannedTargets $Row) $true) } else { $null }

    if ($Row.ErrorText) {
        $Row.RowState = 'Error';    $Row.Status = "Error: $($Row.ErrorText)"
    } elseif ($Row.Selected -and $Row.CanApply) {
        $Row.RowState = 'Change'
        $verb = if ($script:Mode -eq 'Add') { 'exclusion will be added' } else { 'exclusion will be removed' }
        $Row.Status = if ($Row.PreviewOk) { "Preview OK - $verb" } else { "Will change: $verb" }
    } elseif ($Row.Done) {
        $Row.RowState = 'Done';     $Row.Status = 'Updated ' + [char]0x2713
    } elseif ($Row.Selected) {
        $Row.RowState = 'NoChange'
        $Row.Status = switch ($state) {
            'Included' { 'Group is INCLUDED - skipped' }
            'Excluded' { 'Already excluded - nothing to do' }
            default    { 'Not excluded - nothing to remove' }
        }
    } else {
        $Row.RowState = 'Idle';     $Row.Status = ''
    }
}

function Update-AllRows {
    foreach ($r in $script:Rows) { Update-Row $r }
    Update-AssignPanel
    Update-Summary
}

function Update-AssignPanel {
    $sel = Get-HighlightedRows
    $script:AssignItems.Clear()
    if ($sel.Count -eq 0) {
        $ui.lblSel.Text = 'Highlight a row to see its assignments'
        $ui.lblSelInfo.Text = ''
        return
    }
    if ($sel.Count -gt 1) {
        $ui.lblSel.Text = "$($sel.Count) rows highlighted"
        $ui.lblSelInfo.Text = 'Highlight a single row to see its assignments.'
        return
    }
    $r = $sel[0]
    $ui.lblSel.Text = $r.Name
    $ui.lblSelInfo.Text = "$($r.Type)  |  group: $($r.GroupStateText)"

    $cur = @($r.Targets | Where-Object { $null -ne $_ })
    $willChange = $r.Selected -and $r.CanApply
    $planned = if ($willChange) { @(Get-PlannedTargets $r) } else { $cur }
    foreach ($t in $cur) {
        $gone = $willChange -and -not ($planned -contains $t)
        $script:AssignItems.Add((New-Object ExclGui1.AssignItem -Property @{
            Kind = $(if (Test-Exclusion $t) { 'Exclude' } else { 'Include' }); Name = (Get-TargetName $t); Change = $(if ($gone) { 'Remove' } else { '' }) }))
    }
    foreach ($t in $planned) {
        if ($cur -contains $t) { continue }
        $script:AssignItems.Add((New-Object ExclGui1.AssignItem -Property @{
            Kind = $(if (Test-Exclusion $t) { 'Exclude' } else { 'Include' }); Name = (Get-TargetName $t); Change = 'Add' }))
    }
    if ($script:AssignItems.Count -eq 0) { $ui.lblSelInfo.Text += '  |  not assigned' }
}

function Set-Ticked {
    param($Rows, [bool]$Value)
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $r.Selected = $Value
        Update-Row $r
    }
    Update-AssignPanel
    Update-Summary
}

$script:View.Filter = [Predicate[object]]{
    param($r)
    $f = $ui.txtFilter.Text
    if ($f -and $r.Name -notlike "*$f*") { return $false }
    $t = [string]$ui.cmbType.SelectedItem
    if ($t -and $t -ne '(all types)' -and $r.Type -ne $t) { return $false }
    $s = [string]$ui.cmbState.SelectedItem
    if ($s -and $s -ne '(all)' -and $r.GroupStateText -ne $s) { return $false }
    if ($ui.chkTicked.IsChecked -and -not $r.Selected) { return $false }
    if ($ui.chkChanges.IsChecked -and -not ($r.Selected -and $r.CanApply)) { return $false }
    return $true
}

function Update-View { $script:View.Refresh(); Update-Summary }
#endregion

#region Load
function Invoke-Load {
    $grp = ([string]$ui.txtGroup.Text).Trim()
    if (-not $grp) {
        [System.Windows.MessageBox]::Show('Enter the name of the group to exclude.', 'Group', 'OK', 'Warning') | Out-Null
        return
    }

    Set-Busy $true
    $script:Cancellable   = $true
    $ui.btnLoad.Content   = 'Cancel'
    $ui.btnLoad.IsEnabled = $true
    try {
        if (-not $script:WorkerRs) { Write-Log 'Loading the Microsoft Graph module...' }
        try { [void](Get-Worker) }
        catch {
            if ("$_" -like 'Cancelled*') { throw }
            throw "Module Microsoft.Graph.Authentication could not be loaded. Install it with:`nInstall-Module Microsoft.Graph.Authentication -Scope CurrentUser`n`n$(Get-GraphError $_)"
        }

        $tid = ([string]$ui.cmbTenant.Text).Trim()
        # Only reuse a sign-in made in this session, for the same tenant
        $reuse = $script:Connected -and $script:Token -and (-not $tid -or $tid -eq $script:TenantLabel -or $tid -eq $script:TenantInput)
        if ($reuse) {
            Update-TokenIfNeeded
        } else {
            try { [void](Invoke-Worker -Script { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }) } catch {}
            $ui.btnLoad.Content = 'Cancel sign-in'
            $tok = Get-InteractiveToken -Tenant $(if ($tid) { $tid } else { 'organizations' })
            Set-GraphToken $tok
            $ui.btnLoad.Content = 'Cancel'
        }
        $ctx = $script:TokenInfo
        $script:TenantInput = $tid
        $script:Connected   = $true
        $script:TenantLabel = [string]$ctx.TenantId
        $ui.lblConn.Text    = "$($ctx.Account)  |  $($ctx.TenantId)"
        $ui.connDot.Fill    = Get-Brush '#66BB6A'
        Write-Log "Connected as $($ctx.Account) (tenant $($ctx.TenantId))"
        Save-Settings -Group $grp -NameFilter ([string]$ui.txtNameFilter.Text).Trim() -Tenant $(if ($tid) { $tid } else { [string]$ctx.TenantId })

        # Group
        Write-Log "Looking up group '$grp'..."
        $filter = [uri]::EscapeDataString("displayName eq '$($grp.Replace("'", "''"))'")
        $groups = @(Get-GraphAll "$graphV1/groups?`$filter=$filter&`$select=id,displayName")
        if ($groups.Count -ne 1) { throw "Expected exactly 1 group named '$grp', found $($groups.Count)." }
        $script:GroupId    = [string]$groups[0].id
        $script:GroupName  = [string]$groups[0].displayName
        $script:GroupDirty = $false
        $script:GroupNames = @{ $script:GroupId = $script:GroupName }
        Write-Log "Group: $($script:GroupName) ($($script:GroupId))"

        # Profile names first (id + name only = fast), filtered on the name
        $script:Rows.Clear()
        $script:LastApplyOk = $false
        $nf = ([string]$ui.txtNameFilter.Text).Trim()
        $script:LoadFilter = $nf
        $pattern = '*' + ((@($nf -split '\*') | ForEach-Object { [WildcardPattern]::Escape($_) }) -join '*') + '*'
        Write-Log $(if ($nf) { "Loading profiles whose name matches '$nf'..." } else { 'Loading all profiles (no name filter)...' })
        $loaded = New-Object System.Collections.Generic.List[object]
        foreach ($c in $collections) {
            try { $items = @(Get-GraphAll "$graph/$($c.Uri)?`$select=id,$($c.Name)") }
            catch {
                if ("$_" -like 'Cancelled*') { throw }
                Write-Log "Could not retrieve $($c.Type): $(Get-GraphError $_)" 'WARN'; continue
            }
            $hits = @($items | Where-Object { ([string]$_.($c.Name)) -like $pattern })
            foreach ($obj in $hits) {
                $loaded.Add([pscustomobject]@{
                    Name     = ([string]$obj.($c.Name)).Trim()
                    Type     = $c.Type
                    Id       = [string]$obj.id
                    Endpoint = $c.Uri
                    Targets  = [object[]]@()
                })
            }
            Write-Log ("  {0,-18} {1,4} of {2,4} profiles match" -f $c.Type, $hits.Count, $items.Count)
        }

        # Assignments for the matching profiles only, 20 per Graph batch request
        if ($loaded.Count) {
            Write-Log "Loading assignments for $($loaded.Count) profiles..."
            $map = Get-AssignmentsBatch $loaded
            foreach ($l in $loaded) { if ($map.ContainsKey($l.Id)) { $l.Targets = $map[$l.Id] } }
        }

        # Group names for the Included / Excluded columns
        Resolve-GroupNames ([string[]]@($loaded | ForEach-Object { $_.Targets } | ForEach-Object { [string]$_.groupId }))

        foreach ($l in ($loaded | Sort-Object Name)) {
            $r = New-Object ExclGui1.ProfileRow
            $r.Name     = $l.Name
            $r.Type     = $l.Type
            $r.Id       = $l.Id
            $r.Endpoint = $l.Endpoint
            $r.Targets  = $l.Targets
            Update-Row $r
            $script:Rows.Add($r)
        }

        # Fill filters
        $ui.cmbType.Items.Clear(); [void]$ui.cmbType.Items.Add('(all types)')
        foreach ($t in ($script:Rows | ForEach-Object Type | Sort-Object -Unique)) { [void]$ui.cmbType.Items.Add($t) }
        $ui.cmbType.SelectedIndex = 0
        $ui.cmbState.Items.Clear()
        foreach ($s in '(all)', 'Not assigned', 'Excluded', 'Included') { [void]$ui.cmbState.Items.Add($s) }
        $ui.cmbState.SelectedIndex = 0

        Update-View
        Update-AssignPanel
        Write-Log "Done: $($script:Rows.Count) profiles loaded. Next: tick the profiles you want to change (step 2)."
    }
    catch {
        $msg = Get-GraphError $_
        if ($msg -like 'Cancelled*') {
            Write-Log 'Cancelled by user.' 'WARN'
            $script:Rows.Clear()
        } else {
            Write-Log "Load failed: $msg (line $($_.InvocationInfo.ScriptLineNumber))" 'ERROR'
            [System.Windows.MessageBox]::Show("Load failed:`n$msg", 'Error', 'OK', 'Error') | Out-Null
        }
    }
    finally { Stop-Progress; Set-Busy $false }
}
#endregion

#region Apply
function Invoke-Apply {
    $todo = @($script:Rows | Where-Object { $_.Selected -and $_.CanApply })
    if (-not $todo) { [System.Windows.MessageBox]::Show('There are no ticked profiles with changes.', 'Apply changes', 'OK', 'Information') | Out-Null; return }

    $whatIf = [bool]$ui.chkWhatIf.IsChecked
    $verb   = if ($script:Mode -eq 'Add') { 'exclude' } else { 'remove the exclusion of' }

    if (-not $whatIf) {
        $msg  = "You are about to $verb group`n   $($script:GroupName)`n"
        $msg += "on $($todo.Count) configuration profiles in tenant $($script:TenantLabel).`n`n"
        $msg += "All other assignments are kept.`n`nContinue?"
        $ans = [System.Windows.MessageBox]::Show($msg, 'Confirm changes', 'YesNo', 'Warning')
        if ($ans -ne 'Yes') { return }
    }

    Set-Busy $true
    Start-Progress $todo.Count
    $ok = 0; $fail = 0
    $runLog = New-Object System.Collections.Generic.List[object]
    Write-Log ("--- {0}: {1} {2} on {3} profiles ---" -f $(if ($whatIf) { 'PREVIEW' } else { 'APPLY' }), $script:Mode, $script:GroupName, $todo.Count)
    try {
        foreach ($r in $todo) {
            $before  = $r.Excludes
            $planned = @(Get-PlannedTargets $r)
            $after   = Format-Targets $planned $true
            try {
                if ($whatIf) {
                    $r.PreviewOk = $true
                    Update-Row $r -KeepFlags
                    Write-Log "Preview: $($r.Name) [$($r.Type)]  excluded: '$before'  ->  '$after'"
                    $script:LogRows.Add([pscustomobject]@{ Time = Get-Date; Tenant = $script:TenantLabel; Profile = $r.Name; Type = $r.Type; Id = $r.Id; Action = $script:Mode; Group = $script:GroupName; Status = 'Preview'; ExcludedBefore = $before; ExcludedAfter = $after; Included = $r.Includes; Error = '' })
                    Step-Progress
                    continue
                }

                Update-TokenIfNeeded
                $body = @{ assignments = @($planned | ForEach-Object { @{ target = $_ } }) } | ConvertTo-Json -Depth 10
                [void](Invoke-Graph -Method POST -Uri "$graph/$($r.Endpoint)/$($r.Id)/assign" -Body $body)

                # Read back what Intune actually stored
                $fresh = @(Get-GraphAll "$graph/$($r.Endpoint)/$($r.Id)/assignments")
                $r.Targets = [object[]]@($fresh | Where-Object { $_.target } | ForEach-Object { $_.target })
                $r.Done = $true
                Update-Row $r
                $ok++
                Write-Log "OK: $($r.Name) [$($r.Type)]  excluded: '$before'  ->  '$($r.Excludes)'"
                $entry = [pscustomobject]@{ Time = Get-Date; Tenant = $script:TenantLabel; Profile = $r.Name; Type = $r.Type; Id = $r.Id; Action = $script:Mode; Group = $script:GroupName; Status = 'Updated'; ExcludedBefore = $before; ExcludedAfter = $r.Excludes; Included = $r.Includes; Error = '' }
                $script:LogRows.Add($entry); $runLog.Add($entry)
            }
            catch {
                $err = Get-GraphError $_
                if ($err -like 'Cancelled*') { throw }
                $r.ErrorText = $err
                Update-Row $r -KeepFlags
                $fail++
                Write-Log "ERROR: $($r.Name) [$($r.Type)]: $err" 'ERROR'
                $entry = [pscustomobject]@{ Time = Get-Date; Tenant = $script:TenantLabel; Profile = $r.Name; Type = $r.Type; Id = $r.Id; Action = $script:Mode; Group = $script:GroupName; Status = 'Error'; ExcludedBefore = $before; ExcludedAfter = $after; Included = $r.Includes; Error = $err }
                $script:LogRows.Add($entry); $runLog.Add($entry)
            }
            Step-Progress
        }
    }
    finally {
        Stop-Progress
        if ($whatIf) {
            Write-Log 'Preview done - nothing was changed. Untick "Preview only (WhatIf)" in step 4 to apply for real.'
        } else {
            Write-Log "Done: $ok updated, $fail errors."
            $script:LastApplyOk = ($fail -eq 0)
            if ($runLog.Count) {
                try {
                    $safeTenant = ($script:TenantLabel -replace '[^\w\.-]', '_')
                    $file = Join-Path (Get-LogDir) ("GroupExclusionChanges_{0}_{1}.csv" -f $safeTenant, (Get-Date -Format 'yyyyMMdd-HHmmss'))
                    $runLog | Export-Csv -Path $file -Delimiter ';' -NoTypeInformation -Encoding UTF8
                    Write-Log "Audit log saved: $file"
                } catch { Write-Log "Could not save audit log: $($_.Exception.Message)" 'WARN' }
            }
        }
        Set-Busy $false
        Update-AssignPanel
    }
}
#endregion

#region Disconnect
function Invoke-Disconnect {
    $pending = @($script:Rows | Where-Object { $_.Selected -and $_.CanApply }).Count
    if ($pending) {
        $ans = [System.Windows.MessageBox]::Show("There are $pending pending changes that have not been applied.`nThey will be lost. Disconnect anyway?", 'Disconnect', 'YesNo', 'Warning')
        if ($ans -ne 'Yes') { return }
    }
    $old = $script:TenantLabel
    try { Invoke-Worker -Script { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } | Out-Null }
    catch { Write-Log "Disconnect-MgGraph: $(Get-GraphError $_)" 'WARN' }
    if ($old) { Write-Log "Disconnected from tenant $old" } else { Write-Log 'Signed out.' }

    # Clear everything from the previous tenant
    $script:Rows.Clear()
    $script:AssignItems.Clear()
    $script:GroupNames = @{}; $script:GroupId = ''; $script:GroupName = ''; $script:GroupDirty = $false
    $script:Connected = $false; $script:TenantLabel = ''; $script:LastApplyOk = $false
    $script:Token = $null; $script:TokenInfo = $null
    $ui.cmbType.Items.Clear(); $ui.cmbState.Items.Clear()
    $ui.cmbTenant.Text = ''
    $ui.lblConn.Text = 'Not connected'
    $ui.connDot.Fill = Get-Brush '#FFB74D'
    $ui.chkWhatIf.IsChecked = $true
    $ui.btnDisconnect.IsEnabled = $false
    Update-AssignPanel
    Update-Summary
    Write-Log 'Optionally pick another tenant and click "Connect & load".'
}
#endregion

#region Events
function Invoke-Safe {
    param([scriptblock]$Action)
    try { & $Action }
    catch { Write-Log "Error: $($_.Exception.Message) (line $($_.InvocationInfo.ScriptLineNumber))" 'ERROR' }
}

# Step 1
$ui.btnLoad.Add_Click({
    if ($script:Cancellable) { $script:CancelRequested = $true; Write-Log 'Cancelling...' 'WARN'; return }
    Invoke-Safe { Invoke-Load }
})
$ui.btnDisconnect.Add_Click({ Invoke-Safe { Invoke-Disconnect } })
$ui.txtGroup.Add_TextChanged({ Invoke-Safe {
    if ($script:Rows.Count -eq 0) { return }
    $script:GroupDirty = ($ui.txtGroup.Text.Trim() -ne $script:GroupName)
    Update-Summary
} })

# Step 2
$ui.btnSelNotExcl.Add_Click({ Invoke-Safe {
    $rows = @(Get-VisibleRows | Where-Object { $_.GroupState -eq 'None' })
    if (-not $rows) { Write-Log 'No visible profiles without the group.'; return }
    Set-Ticked $rows $true
    Write-Log "$($rows.Count) visible profiles without the group ticked."
} })
$ui.btnSelVisible.Add_Click({ Invoke-Safe { Set-Ticked (Get-VisibleRows) $true } })
$ui.btnSelNone.Add_Click({    Invoke-Safe { Set-Ticked @($script:Rows | Where-Object Selected) $false } })

# Row checkbox (user click only - not fired by scrolling/virtualisation)
$rowClick = [System.Windows.RoutedEventHandler]{
    param($s, $e)
    try {
        $cb = $e.OriginalSource
        if ($cb -isnot [System.Windows.Controls.CheckBox]) { return }
        $row = $cb.DataContext
        if ($row -isnot [ExclGui1.ProfileRow]) { return }
        $row.Selected = ($cb.IsChecked -eq $true)
        Update-Row $row
        Update-AssignPanel
        Update-Summary
    } catch { Write-Log "Error: $($_.Exception.Message)" 'ERROR' }
}
$ui.grid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, $rowClick)

# Header checkbox = tick/untick all visible rows
if ($ui.chkHeaderAll) {
    $ui.chkHeaderAll.Add_Click({ Invoke-Safe { Set-Ticked (Get-VisibleRows) ($ui.chkHeaderAll.IsChecked -eq $true) } })
}

# Spacebar = toggle tick on highlighted rows
$ui.grid.Add_PreviewKeyDown({
    param($s, $e)
    if ($e.Key -ne 'Space') { return }
    $sel = Get-HighlightedRows
    if ($sel.Count -eq 0) { return }
    Set-Ticked $sel (-not $sel[0].Selected)
    $e.Handled = $true
})

# Step 3
$ui.rbAdd.Add_Checked({    Invoke-Safe { $script:Mode = 'Add';    Update-AllRows; Write-Log 'Action: Add exclusion' } })
$ui.rbRemove.Add_Checked({ Invoke-Safe { $script:Mode = 'Remove'; Update-AllRows; Write-Log 'Action: Remove exclusion' } })

# Step 4
$ui.chkWhatIf.Add_Click({ Invoke-Safe { Update-Summary } })
$ui.btnApply.Add_Click({ Invoke-Safe { Invoke-Apply } })
$ui.btnExport.Add_Click({ Invoke-Safe {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV (*.csv)|*.csv'
    $dlg.FileName = "GroupExclusions-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $script:Rows | Select-Object Name, Type, Id, Includes, Excludes, @{ n = 'Group'; e = { $_.GroupStateText } }, Selected, Status |
        Export-Csv -Path $dlg.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
    if ($script:LogRows.Count) {
        $logFile = [IO.Path]::ChangeExtension($dlg.FileName, '.log.csv')
        $script:LogRows | Export-Csv -Path $logFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-Log "Session log exported to $logFile"
    }
    Write-Log "Overview exported to $($dlg.FileName)"
} })

# Highlight -> assignments panel
$ui.grid.Add_SelectionChanged({ Invoke-Safe { Update-AssignPanel } })

# Context menu
$ui.mnuTick.Add_Click({   Invoke-Safe { Set-Ticked (Get-HighlightedRows) $true } })
$ui.mnuUntick.Add_Click({ Invoke-Safe { Set-Ticked (Get-HighlightedRows) $false } })
$ui.mnuCopy.Add_Click({   Invoke-Safe {
    $names = @(Get-HighlightedRows | ForEach-Object Name)
    if ($names) { [System.Windows.Clipboard]::SetText(($names -join [Environment]::NewLine)) }
} })

# Filters
$ui.txtFilter.Add_TextChanged({     Invoke-Safe { Update-View } })
$ui.cmbType.Add_SelectionChanged({  Invoke-Safe { Update-View } })
$ui.cmbState.Add_SelectionChanged({ Invoke-Safe { Update-View } })
$ui.chkTicked.Add_Click({           Invoke-Safe { Update-View } })
$ui.chkChanges.Add_Click({          Invoke-Safe { Update-View } })

$window.Add_Closing({ if ($script:Busy) { $script:CancelRequested = $true } })
$window.Add_Closed({
    # Sign out and clean up the background runspace
    try {
        if ($script:WorkerRs -and -not $script:Busy) {
            $ps = [PowerShell]::Create(); $ps.Runspace = $script:WorkerRs
            [void]$ps.AddScript('Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null')
            [void]$ps.Invoke(); $ps.Dispose()
            $script:WorkerRs.Close()
        }
    } catch {}
})
#endregion

#region Start
$settings = Read-Settings
if ($PSBoundParameters.ContainsKey('GroupName') -or -not $settings.LastGroup) { $ui.txtGroup.Text = $GroupName }
else { $ui.txtGroup.Text = $settings.LastGroup }
if ($PSBoundParameters.ContainsKey('NameFilter') -or $null -eq $settings.LastNameFilter) { $ui.txtNameFilter.Text = $NameFilter }
else { $ui.txtNameFilter.Text = $settings.LastNameFilter }
Update-TenantCombo
if ($TenantId) { $ui.cmbTenant.Text = $TenantId }
Update-AssignPanel
Update-Summary
Write-Log 'Step 1: pick a tenant (optional) and check the group, then click "Connect & load".'
[void]$window.ShowDialog()
#endregion

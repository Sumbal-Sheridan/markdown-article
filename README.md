# Working with Excel in PowerShell

## Automating Excel Reports with the ImportExcel Module

*Written by **Sumbal** – October 2026*

---

### Table of Contents

1. [Introduction](#introduction)
2. [Lab Setup](#lab-setup)
3. [Step-by-Step Commands](#step-by-step-commands)
4. [Results](#results)
5. [Problems I Solved](#problems-i-solved)
6. [Checklist](#checklist)
7. [Conclusion](#conclusion)

---

## Introduction

The **ImportExcel** module lets PowerShell read and write Excel files *without* Microsoft Excel being installed. This is ***very useful*** on servers, where installing Office is not a good idea.

> "If you need to generate reports for work, you must learn this module."
>
> > — from the lab handout

In this article I explain how I used the module to:

- Import data from an Excel file
- Export system information to Excel
- Add new worksheets
- Format tables and create charts

---

## Lab Setup

My lab used **two machines**:

| Role   | Machine            | IP Address       | Tool used       |
| :----- | :----------------: | ---------------: | :-------------- |
| Client | My laptop          | —                | VS Code         |
| Server | DC01 (Server 2022) | `172.24.127.142` | Windows PowerShell |

The server was a Hyper-V virtual machine in the domain `sumbal.powershell.com`.

### Requirements

1. Windows PowerShell 5.1 or newer
2. The ImportExcel module
   1. Installed on the **client**
   2. Installed on the **server**
3. PowerShell remoting enabled
   - `Enable-PSRemoting` on the server
   - `TrustedHosts` set on the client

---

## Step-by-Step Commands

### 1. Install the module

```powershell
Install-Module -Name ImportExcel
```

### 2. Connect to the server and copy the file

```powershell
$cred    = Get-Credential sumbal.powershell.com\Administrator
$session = New-PSSession 172.24.127.142 -Credential $cred
Copy-Item "$env:USERPROFILE\Downloads\Fruits.xlsx" `
          -Destination "C:\Users\Administrator\Documents\Fruits.xlsx" `
          -ToSession $session
Enter-PSSession $session
```

### 3. Import the data

```powershell
Import-Excel "C:\Users\Administrator\Documents\Fruits.xlsx" -OutVariable fruit
$fruit | Get-Member
```

The output showed the data is a `PSCustomObject` with two `NoteProperty` members: `Fruits` and `Quantity`.

### 4. Export, append and format

```powershell
Get-Process | Export-Excel "C:\Users\Administrator\Documents\ProcessInfo.xlsx" -WorksheetName Processes

$people | Export-Excel C:\temp\ExcelDemo.xlsx -Append -WorksheetName "PeopleFormatted" `
          -TableStyle Medium16 -Title "Demo of Table Formatting" -TitleBold
```

### 5. Create a chart

```powershell
$ChartData = New-ExcelChartDefinition -XRange Fruits -YRange Quantity `
             -ChartType ColumnClustered -Title "Total Fruit Sales"
$fruit | Export-Excel C:\temp\ExcelDemo.xlsx -Append -WorksheetName FruitSalesChart `
         -ExcelChartDefinition $ChartData -AutoNameRange -Title "Fruit Sales"
```

---

## Results

The final chart showed the quantity of each fruit:

![Total Fruit Sales chart](images/fruit-chart.png "Chart created with ImportExcel")

| Fruit       | Quantity |
| ----------- | -------: |
| Bananas     | 12       |
| Blueberries | **18**   |
| Apples      | 2        |
| Oranges     | 11       |
| Pears       | 15       |
| Cherries    | 9        |

Blueberries had the highest quantity and Apples had the lowest.

---

## Problems I Solved

Not everything worked the first time. ~~I thought the handout commands would work exactly as written.~~ Here is what I had to fix:

* **Module would not download** – The server could not reach the PowerShell Gallery.
  * Fix: added a DNS forwarder (`8.8.8.8`) and turned on TLS 1.2.
* **Wrong columns** – The handout used columns `1,46,12,6`, but on Server 2022 the correct ones were `1,48,10,5`.
* **Wrong variable names** – The chart example used `$data`, `Item` and `TotalSold`, which did not exist.
  * Fix: used `$fruit`, `Fruits` and `Quantity`.
* **`-Show` failed** – Excel is not installed on the server, so I copied the file back to my laptop to open it.

> **Tip:** Always check the column names with `Get-Member` before writing your commands.

---

## Checklist

- [x] Install ImportExcel on client and server
- [x] Copy Fruits.xlsx to the server
- [x] Import and inspect the data
- [x] Export processes and append a worksheet
- [x] Format a table
- [x] Create a chart
- [ ] Try `Add-ConditionalFormatting` next time

---

## Conclusion

The ImportExcel module makes it easy to build Excel reports from PowerShell, even on a server without Office. The most important lesson was to **check the real output** of each command instead of trusting the instructions blindly.

### Useful Links

- Official module page: <https://www.powershellgallery.com/packages/ImportExcel>
- Source code on [GitHub][importexcel-github]
- Microsoft docs on [PowerShell remoting](https://learn.microsoft.com/powershell/scripting/learn/remoting/running-remote-commands "Running remote commands")

[importexcel-github]: https://github.com/dfinke/ImportExcel

---

<sub>Special characters can be escaped with a backslash, for example \*not italic\* and \# not a heading.</sub>

Line one of an address  
Line two of an address (a line break made with two trailing spaces)

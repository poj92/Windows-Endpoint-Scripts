# Add printer port and printer.
Add-PrinterPort -Name "IP_10.2.0.18" -PrinterHostAddress "10.2.0.18"
Add-Printer -Name "Ricoh Stairwell Printer" -DriverName "PCL6 Driver for Universal Print" -PortName "IP_10.2.0.18"

# Project for subject Architecture of the .NET Technology  
Development of Windows Service which periodically sends emails with manually generated chart as image.

I've used API with forex data. 
- Data are downloaded as JSON and only close prices with dates are parsed (Newtonsoft JSON lib)
- Chart is created using System.Drawings
- Emails are sent every minute. For this, Gmail SMPT server was used

### Requirements:
- Service will download data from some API
- Downloaded data will be visualized by chart. Chart will be sent periodically by email
- It is prohibited to use foreign components / libraries (Chart generated manually as image)

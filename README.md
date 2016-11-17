# Analysis Services (SSAS) on a shoestring

## Setting up a free Analysis Service server

Suppose you want to use Analysis Services but don't have access to SQL Server Enterprise, Business Intelligence or Standard editions (there's an additional twist when it comes to the latter as it can only work in the multidimensional AS mode). Not sure if this was intended, but it turns out that Microsoft was kind enough to package the Analysis Services engine (msmdsrv.exe) with their Power BI Desktop tool which is available free of charge. This essentially gives you everything you need to set up an Analysis Services server - and yes, that includes the tabular model.

In order to get started, do the following:
- Install Power BI Desktop (here I'm assuming you're using the 64-bit version).
- Create the msmdsrv.ini file containing the server properties. You can read about the various options here: https://msdn.microsoft.com/en-us/library/ms174556.aspx. Most importantly, make sure you know the PID (process ID) of the application that will call the Analysis Services engine (i.e. the PrivateProcess property). I wasn't able to get it to work without this. For me, this meant finding the PID of the Command Prompt. Also, set the DeploymentMode setting (multidimensional, tabular, Power Pivot for SharePoint) and pick the port that works for you.
- Start the msmdsrv.exe engine by pointing it at the folder where you put the .ini file:
```
"C:\Program Files\Microsoft Power BI Desktop\bin\msmdsrv.exe" -c -s [YOUR FOLDER (in quotes if path contains spaces)]
```
- Your Analysis Services server is now running at "localhost:port" and can be queried using any of the clients that can talk to an AS server. For example, you can connect to the engine using the (free) SQL Server Management Studio. Note that your choice of the DeploymentMode property determines what you can do with your instance.
- Alternatively, you can simply start the Power BI Desktop (which in turn starts the msmdsrv.exe engine) and determine the randomly-assigned port that the AS engine is using. Then connect to the instance using "localhost:port" as above. However, note that in this case you will not be able to set the location of your instance and its type; it'll be located in a random subfolder in 
```
C:\Users\[USER]\AppData\Local\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces
```
folder and it will always be of the Power Pivot for SharePoint type (i.e. DeploymentMode==1).

Here's an example of the msmdsrv.ini file used to initiatize the AS engine in the tabular mode:
```
<ConfigurationSettings>
   <DataDir>[YOUR FOLDER]</DataDir>
   <TempDir>[YOUR FOLDER]</TempDir>
   <LogDir>[YOUR FOLDER]</LogDir>
   <BackupDir>[YOUR FOLDER]</BackupDir>
   <DeploymentMode>2</DeploymentMode>
   <RecoveryModel>1</RecoveryModel>
   <DisklessModeRequested>0</DisklessModeRequested>
   <CleanDataFolderOnStartup>1</CleanDataFolderOnStartup>
   <AutoSetDefaultInitialCatalog>1</AutoSetDefaultInitialCatalog>
   <Network>
      <Requests>
         <EnableBinaryXML>1</EnableBinaryXML>
         <EnableCompression>1</EnableCompression>
      </Requests>
      <Responses>
         <EnableBinaryXML>1</EnableBinaryXML>
         <EnableCompression>1</EnableCompression>
         <CompressionLevel>9</CompressionLevel>
      </Responses>
      <ListenOnlyOnLocalConnections>1</ListenOnlyOnLocalConnections>
   </Network>
   <Port>[YOUR PORT]</Port>
   <PrivateProcess>[PID]</PrivateProcess>
   <InstanceVisible>0</InstanceVisible>
   <Language>1033</Language>
   <Debug>
      <CallStackInError>0</CallStackInError>
   </Debug>
   <Log>
      <Exception>
         <CrashReportsFolder>[YOUR FOLDER]</CrashReportsFolder>
      </Exception>
      <FlightRecorder>
         <Enabled>0</Enabled>
      </FlightRecorder>
   </Log>
   <AllowedBrowsingFolders>[YOUR FOLDER]</AllowedBrowsingFolders>
   <ResourceGovernance>
      <GovernIMBIScheduler>0</GovernIMBIScheduler>
   </ResourceGovernance>
   <Feature>
      <ManagedCodeEnabled>1</ManagedCodeEnabled>
   </Feature>
   <VertiPaq>
      <EnableDisklessTMImageSave>0</EnableDisklessTMImageSave>
      <EnableProcessingSimplifiedLocks>1</EnableProcessingSimplifiedLocks>
   </VertiPaq>
</ConfigurationSettings>
```

## Next steps - querying the Power Pivot model using Python 2.7

What now? Well, there's all sorts of things you can do. Below is a snippet showing you how to take a Power Pivot model (which is nothing else but an ABF database backup that's packaged inside of an Excel file) and use AMO.NET assembly to restore the model into our new AS instance, then query it using ADOMD.NET. [I'm using the Excel 2013 version of the sample Power Pivot file found here: https://www.microsoft.com/en-us/download/details.aspx?id=102]

Note that this is a POC-type-of code; it just illustrates the approach and doesn't purport to be optimized in any way.

Step 1: Load all modules and .NET assemblies used by the code.
```
import psutil, subprocess, random, os, zipfile, shutil, clr, sys, pandas

def initialSetup(pathPowerBI):
    sys.path.append(pathPowerBI)

    #required Analysis Services assemblies
    clr.AddReference("Microsoft.PowerBI.Amo.Core")
    clr.AddReference("Microsoft.PowerBI.Amo")     
    clr.AddReference("Microsoft.PowerBI.AdomdClient")
    
    global AMO, ADOMD
    import Microsoft.AnalysisServices as AMO
    import Microsoft.AnalysisServices.AdomdClient as ADOMD
```

Step 2: Create a random folder, extract the item.data (or item1.data if dealing with the Excel 2010 version) file from the .xlsx file, append .abf to its name, then start the Command Prompt and determine what PID it was assigned by Windows. Then create the msmdsrv.ini settings file and save it in the random folder created. Finally, start the AS engine, connect to it using AMO.NET and finally restore the backup into it.
```
def restorePowerPivot(excelName, pathTarget, port, pathPowerBI):   
    #create random folder
    os.chdir(pathTarget)
    folder = os.getcwd()+str(random.randrange(10**6,10**7))
    os.mkdir(folder)
    
    #extract PowerPivot model (abf backup)
    archive=zipfile.ZipFile(excelName)
    for member in archive.namelist():
        if ".data" in member:
            filename = os.path.basename(member)
            abfname = os.path.join(folder, filename)+".abf"
            source = archive.open(member)
            target = file(os.path.join(folder, abfname), 'wb')
            shutil.copyfileobj(source, target)
            del target
    archive.close()
    
    #start the cmd.exe process to get its PID
    listPIDpre = [proc for proc in psutil.process_iter()]
    process = subprocess.Popen('cmd.exe /k', stdin=subprocess.PIPE)
    listPIDpost = [proc for proc in psutil.process_iter()]
    pid = [proc for proc in listPIDpost if proc not in listPIDpre if "cmd.exe" in str(proc)][0]
    pid = str(pid).split("=")[1].split(",")[0]
    
    #msmdsrv.ini
    msmdsrvText='''<ConfigurationSettings>
       <DataDir>{0}</DataDir>
       <TempDir>{0}</TempDir>
       <LogDir>{0}</LogDir>
       <BackupDir>{0}</BackupDir>
       <DeploymentMode>2</DeploymentMode>
       <RecoveryModel>1</RecoveryModel>
       <DisklessModeRequested>0</DisklessModeRequested>
       <CleanDataFolderOnStartup>1</CleanDataFolderOnStartup>
       <AutoSetDefaultInitialCatalog>1</AutoSetDefaultInitialCatalog>
       <Network>
          <Requests>
             <EnableBinaryXML>1</EnableBinaryXML>
             <EnableCompression>1</EnableCompression>
          </Requests>
          <Responses>
             <EnableBinaryXML>1</EnableBinaryXML>
             <EnableCompression>1</EnableCompression>
             <CompressionLevel>9</CompressionLevel>
          </Responses>
          <ListenOnlyOnLocalConnections>1</ListenOnlyOnLocalConnections>
       </Network>
       <Port>{1}</Port>
       <PrivateProcess>{2}</PrivateProcess>
       <InstanceVisible>0</InstanceVisible>
       <Language>1033</Language>
       <Debug>
          <CallStackInError>0</CallStackInError>
       </Debug>
       <Log>
          <Exception>
             <CrashReportsFolder>{0}</CrashReportsFolder>
          </Exception>
          <FlightRecorder>
             <Enabled>0</Enabled>
          </FlightRecorder>
       </Log>
       <AllowedBrowsingFolders>{0}</AllowedBrowsingFolders>
       <ResourceGovernance>
          <GovernIMBIScheduler>0</GovernIMBIScheduler>
       </ResourceGovernance>
       <Feature>
          <ManagedCodeEnabled>1</ManagedCodeEnabled>
       </Feature>
       <VertiPaq>
          <EnableDisklessTMImageSave>0</EnableDisklessTMImageSave>
          <EnableProcessingSimplifiedLocks>1</EnableProcessingSimplifiedLocks>
       </VertiPaq>
    </ConfigurationSettings>'''
    
    #save ini file to disk, fill it with required parameters
    msmdsrvini = open(folder+"\\msmdsrv.ini", "w")
    msmdsrvText = msmdsrvText.format(folder, port, pid) #{0},{1},{2}
    msmdsrvini.write(msmdsrvText)
    msmdsrvini.close()
    
    #run AS engine inside the cmd.exe process
    initString = "\"{0}\\msmdsrv.exe\" -c -s \"{1}\""
    initString = initString.format(pathPowerBI.replace("/","\\"),folder)
    process.stdin.write(initString + " \n")
    
    #connect to the AS instance from Python
    AMOServer = AMO.Server()
    AMOServer.Connect("localhost:{0}".format(port))
    
    #restore database from PowerPivot abf backup, disconnect
    AMORestoreInfo=AMO.RestoreInfo(os.path.join(folder, abfname))
    AMOServer.Restore(AMORestoreInfo)
    AMOServer.Disconnect()
    
    return process
```

Step 3: Use ADOMD.NET assembly to query the restored database and write the results to a Pandas dataframe.
```   
def runQuery(query,port,flag):
    #ADOMD assembly
    ADOMDConn=ADOMD.AdomdConnection("Data Source=localhost:{0}".format(port))
    ADOMDConn.Open()
    ADOMDCommand=ADOMDConn.CreateCommand() 
    ADOMDCommand.CommandText = query
    
    #read data in via AdomdDataReader object
    DataReader = ADOMDCommand.ExecuteReader()
    
    #get metadata, number of columns
    SchemaTable=DataReader.GetSchemaTable()
    numCol = SchemaTable.Rows.Count #same as DataReader.FieldCount
    
    #get column names
    columnNames = []
    for i in range(numCol):
        columnNames.append(str(SchemaTable.Rows[i][0]))
    
    #fill with data
    data = []
    while DataReader.Read()==True:
        row=[]
        for j in range(numCol):
            try:
                row.append(DataReader[j].ToString())
            except:
                row.append(DataReader[j])
        data.append(row)
    df = pandas.DataFrame(data)
    df.columns = columnNames 
    
    if flag==0:
        DataReader.Close()
        ADOMDConn.Close()
	
        return df     
    else:	
        #metadata table
        metadataColumnNames = []
        for j in range(SchemaTable.Columns.Count):
            metadataColumnNames.append(SchemaTable.Columns[j].ToString())
        metadata=[]
        for i in range(numCol):
            row=[]
            for j in range(SchemaTable.Columns.Count):
                try:
                    row.append(SchemaTable.Rows[i][j].ToString())
                except:
                    row.append(SchemaTable.Rows[i][j])
            metadata.append(row)
        metadf = pandas.DataFrame(metadata)
        metadf.columns = metadataColumnNames
		
        DataReader.Close()
        ADOMDConn.Close()
	
        return df, metadf
```

Step 4: Terminate the session.
```
def endSession(process):
    #terminate cmd.exe
    process.terminate()
    print "Session terminated."
```

### Example

If you download the sample Power Pivot file from Microsoft's website, you can test everything by appropriately modifying the following couple of lines of code:
```
pathPowerBI="C:/Program Files/Microsoft Power BI Desktop/bin"
initialSetup(pathPowerBI)
session = restorePowerPivot("D:/Downloads/PowerPivotTutorialSample.xlsx", "D:/", 60000, pathPowerBI)
df = runQuery("EVALUATE dbo_DimProduct",60000,0)
df, metadf = runQuery("EVALUATE dbo_DimProduct",60000,1)
endSession(session)
```

## Final thoughts
If you want to restore the ABF file into the Power Pivot for Sharepoint type of AS instance, you need to do the following:
```
clr.AddReference("System.IO") # mscorlib.dll
import System.IO as SIO
Stream = SIO.FileStream(target, 3) # 3 = FileMode.Open
AMOImageLoadInfo=AMO.ImageLoadInfo("SomeDatabaseName","SomeDatabaseID", Stream, 0) # 0 = ReadWrite mode
AMOServer.ImageLoad(AMOImageLoadInfo)
AMOServer.Refresh()
Stream.Close()
```

And if you want to execute some arbitrary XMLA code, you can do the following:
```
AMOServer.Execute(YourXMLAcodeString)
```

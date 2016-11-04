cls
#If necessary add SharePoint modules
if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null) {
Add-PSSnapin Microsoft.SharePoint.PowerShell;
}

#Get current date
$todaysDate = Get-Date;
$todayYear = $todaysDate.Year;
$todayMonth = $todaysDate.Month;
$todayDay = $todaysDate.Day;

#User defined variables
$outputXmlFilePath="C:\OutBox\PanelA.xml";
$webURL = "http://";
$XSLPropText = 'type="text/xsl" href="status_board.xsl"'
$mfgTrackingListName = "Dev MFG Tracking";
$actionItemTrackingListName = "DevAction Item Tracking";
$masterTimelineListName = "DevMaster Timeline";
#End user defined variables

#Non user defined variables
$itemArray  = @();
#End non user defined variables

$spWeb = Get-SPWeb $webURL;

#MFG Tracking Items
$spMFGTrackingList = $spWeb.Lists[$mfgTrackingListName];
$spMFGTrackingItems = $spMFGTrackingList.GetItems();

#Action Items Tracking Items
$spActionItemTrackingList = $spWeb.Lists[$actionItemTrackingListName];
$spACtionItemTrackingItems = $spActionItemTrackingList.GetItems()

#Master Timeline Items
$spMasterTimelineList = $spWeb.Lists[$masterTimelineListName];
$spMasterTimelineItems = $spMasterTimelineList.GetItems()

#Get IDs of MFG Tracking list items
$spMFGTrackingItems | ForEach-Object {
    $itemProgess = $_['DevMFG Progression'];
    if ($itemProgess -ne "Parked")
    {
      $itemArray += $_.ID;
    }
}
#End getting IDs
$listIndex = 0;
$arrayIndex = 0;

while ($itemArray[$listIndex])
{
#Start to create XML
[System.Xml.XmlTextWriter]$xml = New-Object 'System.Xml.XmlTextWriter' $outputXmlFilePath, ([Text.Encoding]::UTF8);
$xml.Formatting = "indented";
$xml.Indentation = 4;
$xml.WriteStartDocument();
$xml.WriteProcessingInstruction("xml-stylesheet", $XSLPropText)

#Create root node
$xml.WriteStartElement('board');

$currentIndex = 0;

#Loop through the MFG Tracking List
while ($currentIndex -ne 2 -and $itemArray[$arrayIndex] -ne $null)
{
$spMFGTrackingItem = $spMFGTrackingItems.GetItemById($itemArray[$arrayIndex]);

#Pull the procession from the current item
$itemProgess = $spMFGTrackingItem['DevMFG Progression'];
#Pull the SKU from the current item
$mfgSKU = $spMFGTrackingItem['DevSKU'];

#Create the second level node
$xml.WriteStartElement('store');

#Create status node
  if ($spMFGTrackingItem['DevMFG SubState'] -eq "parked")
    {      
      $xml.WriteElementString("status","pause");
    }
  elseif ($itemProgess -eq "7-Complete")
    {
      $xml.WriteElementString("status","stop");
    }
  Else
    {
      $xml.WriteElementString("status","go");
    }
#End status node

#Add current item SKU to XML
$xml.WriteElementString("sku",$spMFGTrackingItem['DevSKU']);

#Create the Lab node
$xml.WriteStartElement('lab');
$xml.WriteElementString("factory",$spMFGTrackingItem['DevFactory']);
#Creating Lab ID node
$lab = $spMFGTrackingItem['DevLab'];

if ((!$lab -eq "Warehouse") -or (!$lab -eq ""))
    {      
       $lab = $lab.TrimStart("Lab-");
       $xml.WriteElementString("lab_id",$lab);
    }
$xml.WriteElementString("ip",$spMFGTrackingItem['DevFactory IP']);
$xml.WriteElementString("state",$spMFGTrackingItem['DevMFG Progression']);
#Softpacs
if ($spMFGTrackingItem['DevSoftpacs'] -EQ "FALSE")
    {      
      $xml.WriteElementString("softpacs","Yes");
    }
  Else
    {
      $xml.WriteElementString("softpacs","No");
    }

$xml.WriteEndElement();
#End Lab node

#Create MBO node
$xml.WriteStartElement('mbo');
#Create network node
  if (!$spMFGTrackingItem['DevMBO'])
    {      
      $xml.WriteElementString("network","Vendor");
    }
  Else
    {
      $xml.WriteElementString("network","Facility");
    }
$xml.WriteElementString("public_ip",$spMFGTrackingItem['DevMBO:DevStatic IP']);
$xml.WriteElementString("store_ip",$spMFGTrackingItem['DEVFacility Subnet']);
$xml.WriteElementString("agent",$spMFGTrackingItem['DevAgent']);
$xml.WriteEndElement();
#End MBO node

$numberActionItems=0;
#Loop through the Action Item Tracking List
$spACtionItemTrackingItems | ForEach-Object {      
  #Pull the SKU from the current item    
  $actionItemSKU = $spMFGTrackingItem['DevSKU'];
  #Increment $numberActionItems if MFG Tracking Item SKU matches the Action Item Tracking Item SKU
  if ($mfgSKU -eq $actionItemSKU)
    {
      $numberActionItems++; 
    }    
}

$xml.WriteStartElement('dates');

#Loop through the Master Timeline List
$spMasterTimelineItems | ForEach-Object {     
  $masterTimelineItemSKU =  $_['DevSKU'];
  
  $masterTimelineItemDateType = $_['Date Type'];
  
  #Get due date from item in SharePoint
  $masterTimelineItemDate = Get-Date $_['Due Date'];   
  $masterTimelineItemYear = $masterTimelineItemDate.Year;  
  $masterTimelineItemMonth = $masterTimelineItemDate.Month;  
  $masterTimelineItemDay = $masterTimelineItemDate.Day;  

  if ($masterTimelineItemSKU -like "*$mfgSKU")
    {      
      #Check to see type of date it is
      if ($masterTimelineItemDateType -eq "Strategies")
        {
          $xml.WriteStartElement('strategies');
          $xml.WriteElementString("date",$_['Due Date']);

          $masterTimelineItemTaskStatus = $_['Task Status'];

          if ($masterTimelineItemTaskStatus -eq "Completed")
            {
              $xml.WriteElementString("deadline_met","yes");
            }                       
          else
          {          
            if (($todayYear -ge $masterTimelineItemYear) -and ($todayMonth -ge $masterTimelineItemMonth) -and ($todayDay -ge $masterTimelineItemDay))
              {
                if (($todayYear -eq $masterTimelineItemYear) -and ($todayMonth -eq $masterTimelineItemMonth) -and ($todayDay -eq $masterTimelineItemDay))
                  {
                    #Last day before deadline is not met
                  }
                else
                  {
                    $xml.WriteElementString("deadline_met","no");
                  }
              }
          }
          #strategies
          $xml.WriteEndElement();
        }
      elseif ($masterTimelineItemDateType -eq "System Integration")
        {      
          $xml.WriteStartElement('system_int');
          $xml.WriteElementString("date",$_['Due Date']);
          
          $masterTimelineItemTaskStatus = $_['Task Status'];
          
          if ($masterTimelineItemTaskStatus -eq "Completed")
            {
              $xml.WriteElementString("deadline_met","yes");
            }                       
          else
          {
            if (($todayYear -ge $masterTimelineItemYear) -and ($todayMonth -ge $masterTimelineItemMonth) -and ($todayDay -ge $masterTimelineItemDay))
              {
                if (($todayYear -eq $masterTimelineItemYear) -and ($todayMonth -eq $masterTimelineItemMonth) -and ($todayDay -eq $masterTimelineItemDay))
                  {
                    #Last day before deadline is not met
                  }
                else
                  {
                    $xml.WriteElementString("deadline_met","no");
                  }
              }
          }
          #system_int
          $xml.WriteEndElement();          
        }
      elseif ($masterTimelineItemDateType -eq "Commissioning")
        {      
          $xml.WriteStartElement('commissioning');
          $xml.WriteElementString("date",$_['Due Date']);

          $masterTimelineItemTaskStatus = $_['Task Status'];

          if ($masterTimelineItemTaskStatus -eq "Completed")
            {
              $xml.WriteElementString("deadline_met","yes");
            }                       
          else
          {
            if (($todayYear -ge $masterTimelineItemYear) -and ($todayMonth -ge $masterTimelineItemMonth) -and ($todayDay -ge $masterTimelineItemDay))
              {
                if (($todayYear -eq $masterTimelineItemYear) -and ($todayMonth -eq $masterTimelineItemMonth) -and ($todayDay -eq $masterTimelineItemDay))
                  {
                    #Last day before deadline is not met
                  }
                else
                  {
                    $xml.WriteElementString("deadline_met","no");
                  }
              }
          } 
          #commissioning
          $xml.WriteEndElement();          
        }
      elseif ($masterTimelineItemDateType -eq "Soft Open")
        {      
          $xml.WriteStartElement('soft_open');
          $xml.WriteElementString("date",$_['Due Date']);   
          
          $masterTimelineItemTaskStatus = $_['Task Status'];

          if ($masterTimelineItemTaskStatus -eq "Completed")
            {
              $xml.WriteElementString("deadline_met","yes");
            }                       
          else
          {
            if (($todayYear -ge $masterTimelineItemYear) -and ($todayMonth -ge $masterTimelineItemMonth) -and ($todayDay -ge $masterTimelineItemDay))
              {
                if (($todayYear -eq $masterTimelineItemYear) -and ($todayMonth -eq $masterTimelineItemMonth) -and ($todayDay -eq $masterTimelineItemDay))
                  {
                    #Last day before deadline is not met
                  }
                else
                  {
                    $xml.WriteElementString("deadline_met","no");
                  }
              }
          }
          #soft_open
          $xml.WriteEndElement();
        }
      elseif ($masterTimelineItemDateType -eq "Deadline")
        {      
          $xml.WriteStartElement('deadline');
          $xml.WriteElementString("date",$_['Due Date']); 
          
          $masterTimelineItemTaskStatus = $_['Task Status'];

          if ($masterTimelineItemTaskStatus -eq "Completed")
            {
              $xml.WriteElementString("deadline_met","yes");
            }                       
          else
          {
            if (($todayYear -ge $masterTimelineItemYear) -and ($todayMonth -ge $masterTimelineItemMonth) -and ($todayDay -ge $masterTimelineItemDay))
              {
                if (($todayYear -eq $masterTimelineItemYear) -and ($todayMonth -eq $masterTimelineItemMonth) -and ($todayDay -eq $masterTimelineItemDay))
                  {
                    #Last day before deadline is not met
                  }
                else
                  {
                    $xml.WriteElementString("deadline_met","no");
                  }
              }
          }     
          #deadline
          $xml.WriteEndElement();
        }
    }    
 }
 #dates
 $xml.WriteEndElement();
 
 #Create action_items node
$xml.WriteStartElement('action_items');
#Insert into the open node that number of action items associated with the MFG Tracking Item
$xml.WriteElementString("open",$numberActionItems);

#Create the build node
$xml.WriteElementString("build",$spMFGTrackingItem['DevBuild']);
#action_items
$xml.WriteEndElement();
#End action_items node

#Created comments node
########################################comments are inlcuding formattting text######################
$xml.WriteStartElement('comments');
$xml.WriteElementString("comment",$spMFGTrackingItem['DevDaily Summary']);
$xml.WriteEndElement();
#End comments node

 #store
 $xml.WriteEndElement(); 
 $currentIndex++;
 $arrayIndex++;
 $listIndex++;
}
#board
$xml.WriteEndElement();

#Cleanup
$xml.Flush();
$xml.Close();

#$listIndex++;

switch ($listIndex) 
    { 
        0 {$outputXmlFilePath="C:\OutBox\PanelA.xml";}         
        1 {$outputXmlFilePath="C:\OutBox\PanelA.xml";} 
        2 {$outputXmlFilePath="C:\OutBox\PanelB.xml";} 
        3 {$outputXmlFilePath="C:\OutBox\PanelB.xml";} 
        4 {$outputXmlFilePath="C:\OutBox\PanelC.xml";}         
        5 {$outputXmlFilePath="C:\OutBox\PanelC.xml";}         
        6 {$outputXmlFilePath="C:\OutBox\PanelD.xml";}         
        7 {$outputXmlFilePath="C:\OutBox\PanelD.xml";}         
        8 {$outputXmlFilePath="C:\OutBox\PanelE.xml";}         
        9 {$outputXmlFilePath="C:\OutBox\PanelE.xml";}         
        default {"NA"}
    }
#Clear-Variable -Name xml;
}
$spWeb.Dispose();
"Completed"
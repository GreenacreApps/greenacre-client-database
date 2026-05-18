Greenacre Client Database - SharePoint lists working version

This version writes to the four SharePoint lists created for the client app:

  GCE_Clients
  GCE_ClientContacts
  GCE_Sites
  GCE_Assets

It no longer uses GCE_ClientRecords for SharePoint saving.

Required columns
----------------
GCE_Clients:
  Title
  BillingAddress
  ClientNotes

GCE_ClientContacts:
  Title
  Client        lookup to GCE_Clients
  Role
  Email
  Phone

GCE_Sites:
  Title
  Client        lookup to GCE_Clients
  SiteReference
  SiteAddress

GCE_Assets:
  Title
  Site          lookup to GCE_Sites
  AssetReference
  AssetLocation
  FlowRate
  FanModel
  FanSerial
  FlowSensorModel
  FlowSensorSerial
  FlowSetpointPercent
  CarbonMediaType
  CarbonMediaSlNumber
  CarbonMediaVolume
  CarbonPressureDrop
  CarbonHighSetpoint
  CarbonHighHighSetpoint
  CarbonTempProbeModel
  CarbonTempProbeSerial
  BioMediaType
  BioMediaVolume

Files
-----
GreenacreClientDatabase-SharePoint-Lists.html
  Single-file version for opening/uploading.

index.html + app.js + styles.css
  Editable source version.

Workflow
--------
Client Register -> Site Register -> Asset Register -> Asset Info

Save buttons are provided for:
  Save Client
  Save Site
  Save Asset

Asset reference and asset location use dropdown-style fields based on existing saved assets.

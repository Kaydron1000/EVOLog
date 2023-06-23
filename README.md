Currently in development

# EVOLog - Evocation Logging
Evoke Verbose Output Logging - Channel multiple log outputs with EVOLog

This project is a VBA logging project that was inspired by Serilog style logging. VBA is an obtuse language that is easily avaiable to those using Microsoft Office. This project aims at providing logging capablitiy that can be channeled out to multiple conduits. 

EVOLog contains a Logger that can add conduits that can channel log messages to different conduits (i.e. File, counter, textbox, etc.)

## cEVOLogger Class:
EVOLogger's base class is the cEVOLogger. This class allows channeling log messages to multiple conduits. Once the class is instantiated conduits must be added to channel log messages out to the conduits destination. Conduits implement the ILogConduit interface to channel log messages to any destination (i.e. Text File,Text Box, another cEvoLogger).

## ILogConduit Interface:
EVOLogger has the capablity to interface with any ILogConduit created. This conduit definition provides the neccessary properties and subroutines for cEVOLogger class to utilze a conduit. Conduits provide a way to channel a log message to its appropriate destination (i.e. Text File,Text Box).

cEvoLogger
-LoggerName
-LoggingLevelNames
-BatchOutput
-BatchSetCount
-----------------------------
-CounditsCount
-GetCouduitNames
-AddConduit
-GetConduit
-RemoveConduit
-ClearConduits
-----------------------------
-LogArtifact
-LogArtifactObject
-----------------------------
-Init

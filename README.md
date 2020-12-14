# csv-to-publishednodes

Converts CSV files to published nodes json file to be used with Azure Industrial IoT OPC Publisher module.

USAGE: Convert-CSVToPublishedNodes.ps1  [OPTIONS]

SAMPLE: Convert-CSVToPublishedNodes.ps1  -InputFileName sample.csv

OPTIONS:
          -h, Help : (Optional)  Display this screen.
          -InputFileName : (Mandatory) Path to the input file. Needs to be in CSV format.
          -OutputFileName : (Optional)  Path to the output file. Default: Same name as input file, with JSON extension.
          -Delimiter : (Optional)  Column delimiter in CSV file. Default: ','

* InputFileName includes a line for each distinct OPC Node (OpcNodes_Id or OpcNodes_ExpandedNodeId).

* Only mandatory columns required to exist in the file (EndpointUrl and (OpcNodes_Id or OpcNodes_ExpandedNodeId))

* Ordering of columns is arbitrary. Only column names have to match below column structure. 

* Server related properties (ones that don't start with 'OPCNodes_') repeat for each OPC Node on that server. 

* If server related properties (ones that don't start with 'OPCNodes_') are null for a given line, last known properties from previous lines are used.

* InputFileName column structure is as follows:

    | Column                         | Type                   | Description                                                  |
    | ------------------------------ | ---------------------- | ------------------------------------------------------------ |
    | EndpointUrl                    | [mandatory], [string]  | URL of OPC UA Server in format "opc.tcp://<your_opcua_server>:<your_opcua_server_port>/<your_opcua_server_path>". |
    | UseSecurity                    | optional] , [boolean]  | Allows to access the endpoint with SecurityPolicy.None when set to 'false' (no signing and encryption applied to the OPC UA communication), default is true |
    | OpcAuthenticationMode          | [optional] , [string]  | "Anonymous" or "UsernamePassword", default is "Anonymous"    |
    | Username                       | [optional] , [string]  | Valid only if "OpcAuthenticationMode": "UsernamePassword"    |
    | Password                       | [optional] , [string]  | Valid only if "OpcAuthenticationMode": "UsernamePassword"    |
    | OpcNodes_Id                    | [mandatory], [string]  | OPC node to publish in either NodeId format (contains "ns=", e.g. "ns=3;i=1234") or ExpandedNodeId format (contains "nsu=", e.g. "nsu=http://mycompany.com/UA/Data;i=1234").<br />Only one of "Id" or "ExpandedNodeID" is mandatory. |
    | OpcNodes_ExpandedNodeId        | [mandatory], [string]  | OPC node to publish in either NodeId format (contains "ns=", e.g. "ns=3;i=1234") or ExpandedNodeId format (contains "nsu=", e.g. "nsu=http://mycompany.com/UA/Data;i=1234").<br />Included for backward compatibility. |
    | OpcNodes_OpcSamplingInterval   | [optional] , [int]     | Sampling interval OPC Publisher requests the server to sample the node value. The value is in milliseconds. |
    | OpcNodes_OpcPublishingInterval | [optional] , [int]     | Subscription will publish node value with this interval, it will only be published if the value has changed. The value is in milliseconds. |
    | OpcNodes_DisplayName           | [optional] , [string]  | Display Name for Node. This value overrides DisplayName values fetched from server with -fd=true switch. |
    | OpcNodes_HeartbeatInterval     | [optional] , [int]     | If set, the last value will be sent again with an updated SourceTimestamp value after the given interval. The value is in milliseconds. |
    | OpcNodes_SkipFirst             | [optional] , [boolean] | When true, first event will not generate a telemetry event, this is useful when publishing a large amount of data to prevent a event flood at startup of OPC Publisher. |

    

[Microsoft Graph SDK](README.md) / streamParameters

# streamParameters

Streaming parameter utilities for workbook operations.

## Variables

### streamHighWaterMark

> `const` **streamHighWaterMark**: `number`

Defined in: [src/services/streamParameters.ts:11](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/streamParameters.ts#L11)

Maximum amount of data (in bytes) to buffer in memory for streaming operations.
Used as the highWaterMark for streams to control memory usage.

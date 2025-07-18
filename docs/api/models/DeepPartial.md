[Microsoft Graph SDK](../README.md) / models/DeepPartial

# models/DeepPartial

## Type Aliases

### DeepPartial\<T\>

> **DeepPartial**\<`T`\> = `{ [P in keyof T]?: T[P] extends object ? DeepPartial<T[P]> : T[P] }`

Defined in: [src/models/DeepPartial.ts:1](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/DeepPartial.ts#L1)

#### Type Parameters

| Type Parameter |
| ------ |
| `T` |

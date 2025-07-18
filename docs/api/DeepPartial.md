[Microsoft Graph SDK](README.md) / DeepPartial

# DeepPartial

Making all properties of a type optional, recursively.

## Type Aliases

### DeepPartial\<T\>

> **DeepPartial**\<`T`\> = `{ [P in keyof T]?: T[P] extends object ? DeepPartial<T[P]> : T[P] }`

Defined in: [src/models/DeepPartial.ts:13](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/DeepPartial.ts#L13)

Makes all properties of a type optional, recursively.
Useful for partial updates or patch operations on deeply nested objects.

#### Type Parameters

| Type Parameter | Description |
| ------ | ------ |
| `T` | The type to make deeply partial. |

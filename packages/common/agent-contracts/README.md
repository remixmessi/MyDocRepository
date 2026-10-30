# @mydoc/agent-contracts

Portable Agent protocol v1: messages, complete tool calls, lifecycle events,
provider/tool/event ports and bounded run limits. Exports TypeScript sources like
other MyDoc common packages. No Node, React, storage or SDK dependencies.

The protocol includes interaction requests, source references and changeset types;
persistent sessions, approval binding/expiry, source authorization and resumable
session cursors are not implemented. Consumers must not
interpret its TypeScript interfaces as runtime validation of untrusted transports.
`validateCall` validates complete tool identities, strict JSON arguments, depth and
serialized size. `validateTools` bounds catalog descriptions and schemas;
resource-specific argument validation and authorization belong to ToolGateway.

# Graph2AutoTask
Simple application utilizing the graph api to check an email account and then use the autotask api v1.6 to create a ticket.

Supports Ticket Creation uses some Fuzzy customer matching(may not be the best but easy to improve).
Supports Ticket Note Creation
Supports Attachment Creation
Supports Resource Impersonation
Supports dropping attachment images under width and height constraints.
Supports resizing images if over 59000000 bytes(used as AT API limits us to 6MB)
Supports compressing attachments if over 59000000 bytes(used as AT API limits us to 6MB)
Supports adding "[internal]" to the start of the email message which will then flag the corresponding ticket note and/or attachments as internal only.

Not all but a lot of the calls are in a queue manager, the queue manager works in a mostly fifo order, however it will thread itself up
depending on the number of cores in the system.

At also limits how many attachments/data can be sent within a 5 minute time period this is handled in the queue via exceptions/retries.

It's used in production for a medium sized MSP with very few issues.

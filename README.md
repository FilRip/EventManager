# EventManager

The class contains methods to :
- Retreive a list of handler of an event, or a list of event (including WinForm and WPF event)
- Remove all handlers of an event, or a list of event
- Start a BeginInvoke on an event where there is more than one handler
- Copy a list of handlers of an event (or list of events) from an object to another (useful when replace object by another one, no need to reinject all handlers yourself)

Nb : Remove and Copy contains methods to do it on objects included (in fields or properties) in the object copied (or remove handler)<br>
Contains a method to get the Field (who store the handlers) of an event.<br>
Compatible with WinForm & WPF

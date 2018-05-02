Options for catching key combinations in office are suprisingly limited

1. Office javascript API does not support key bindings, but the feature is planned
  - https://officespdev.uservoice.com/forums/224641-feature-requests-and-feedback/suggestions/14060244-enable-binding-keyboard-shortcuts-to-a-button-in
2. Macros can bind to key combos
  - https://stackoverflow.com/questions/22855021/change-keyboard-shortcut-by-add-in-for-office 
  - Option killer: macros can interop with plugins, but must be manually installed
    https://social.msdn.microsoft.com/Forums/vstudio/en-US/2e82df65-f192-4c25-a1ff-ca8a541e989e/can-i-have-a-macro-with-vsto?forum=vsto
3. OnKey appears to only be able to trigger macros and built in functionality
  - https://stackoverflow.com/questions/2635463/i-have-a-vsto-application-as-an-add-in-to-ms-word-and-i-want-to-set-keyboard-sho
4. Global keyboard hook for process
  - https://blogs.msdn.microsoft.com/vsod/2010/04/08/using-shortcut-keys-to-call-a-function-in-an-office-add-in/
  - existing libraries
   - https://github.com/gmamaladze/globalmousekeyhook
   - https://github.com/factormystic/FMUtils.KeyboardHook#readme

Winner (for now): Global keyboard hook

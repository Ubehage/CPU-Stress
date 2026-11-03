Mouse handling analysis and proposed interface

This proposal centralizes pointer tracking, capture ownership, and notification delivery while leaving drawing, keyboard activation, checkbox values, and CPU actions in their controls. Start with `Button`, then migrate `CheckBox` and `CPUView` after the event contract is verified.

This is the original source review, before implementation and runtime validation. Findings below distinguish directly traceable paths from risks requiring reproduction. Source locations refer to the baseline files as reviewed on 2026-11-03. See [the implementation notes](MouseHandling-Implementation.md) for the completed Button, CheckBox, and CPUView migrations, build instructions, and validation status.

**What the current code does**

| Area | Existing technique | What should remain local |
| --- | --- | --- |
| `Button.ctl:316`, `:399` | `MouseMove` starts hover and calls `SetCapture`; cursor position, a cached screen rectangle, and `WindowFromPoint` detect departure. Left down records a press; release inside raises `Click`. | Colors, caption, drawing, Space/Return activation, public `Click`. |
| `CheckBox.ctl:209`, `:272` | Nearly the same hover/capture/press sequence as Button. A successful release toggles the value and raises `Click`. | Value changes, layout, drawing, public `Click`. |
| `CPUView.ctl:525`, `:544`, `:599` | Captures for hover, converts the current cursor position to client twips, and searches CPU rectangles. Remembers the pressed index and compares it with the release index. | CPU rectangles, region hit testing, partial redraws, public indexed `Click`. |
| `modSystem.bas:184`, `:192` | Shared rectangle/window tests. The `MouseIsDown` argument deliberately bypasses outside-window rejection during a press. | Generic geometry may move into the exportable component; application helpers need not move wholesale. |

Here, “hover” means immediate highlighting while over a target. It is not timed dwell. All three controls keep the press active while dragging outside; Button and CheckBox also keep their hover flag true during that drag. CPUView independently tracks the region under the pointer while retaining the original pressed region.

`Frame` and `StatusBar` contain no mouse event handlers. `CPUTimer` already demonstrates a BAS callback dispatching to class instances, but its object-pointer mechanism is not needed for the proposed recipient registry.

**Findings to carry into the design**

| Finding | Evidence and implication |
| --- | --- |
| Capture is used for passive hover. | Every `StartHover` calls `SetCapture`. Only one window can own capture, and delivery across threads is restricted when no button is down. Therefore capture is an unreliable general leave detector; test fast exits and movement into another application. See [SetCapture](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setcapture). |
| Capture ownership can become stale. | Each control sets `m_IsCapturing = True` without checking current ownership. No `GetCapture`, `WM_CAPTURECHANGED`, or `WM_CANCELMODE` handling exists. `EndCapture` trusts that local flag, so it can release capture now belonging to another control on the thread. |
| Disabling during a press skips explicit release. | `Button.ctl:114` and `CheckBox.ctl:73` call `EndHover` before clearing `m_MouseIsDown`. `EndHover` returns while that flag is true; subsequent mouse handlers return because the control is disabled. This path leaves the local capture flag set and does not explicitly release capture. |
| Geometry can become stale. | Button and CheckBox cache screen rectangles when hover starts with a refresh. Button also updates on resize. Neither consistently updates on parent movement during an existing hover; CheckBox resize does not directly refresh that cache. Reproduce movement/layout changes beneath a stationary pointer. |
| CPUView can emit an invalid region. | `GetCPUIndexFromMousePos` returns zero for empty space. Down records zero, and up accepts equality without requiring a positive index (`CPUView.ctl:616`). Thus an empty-space press/release can raise `Click(0)`. `frmMain.frm:409` accepts the index without validation; index zero also identifies the main shared-memory slot. This is a directly traceable path; downstream effects have not been exercised. |
| CPUView release validation is weaker. | Its release path checks geometry alone, without repeating window visibility/occlusion validation. After `RaiseEvent Click`, it resumes hovering without the `CanInteractNow` check used by Button and CheckBox. Test handlers that hide, disable, or unload their source. |
| Lifecycle cleanup is incomplete. | Button and CheckBox have no termination handlers. CPUView termination stops its update timer and closes PDH, but does not reset mouse state. No explicit mouse cancellation handles hiding or application deactivation. Window destruction may release OS capture, but a future registry must also drop its references and registrations. |

The existing pressed appearance outside a target is a behavior choice, not automatically a defect. Preserve it initially; track actual pointer location separately so controls can later choose a different appearance. Also standardize rectangle edges: existing tests include right/bottom edges, whereas the proposed client-pixel contract uses exclusive right/bottom boundaries.

**Proposed component and interface**

Use a small source package: `modMouse.bas` for the registry, Windows integration, and state transitions; `IMouseSink.cls` for a typed recipient contract. Keep API declarations and types self-contained, with no dependency on CPU-Stress globals, controls, or timers. Initial scope is windowed VB6 controls on one UI thread, with one recipient per window. One recipient may register multiple windows.

These are proposed signatures, not implemented code:

```vb
' modMouse.bas
Public Function MouseRegister(ByVal hWnd As Long, _
    ByVal Sink As IMouseSink) As Long
Public Sub MouseUnregister(ByVal Token As Long)
Public Sub MouseSetEnabled(ByVal Token As Long, ByVal Enabled As Boolean)
Public Sub MouseRefresh(ByVal Token As Long)
Public Sub MouseCancel(ByVal Token As Long, ByVal Reason As MouseCancelReason)
Public Sub MouseShutdown()

' IMouseSink.cls: implemented by a control/class or a module adapter
Public Function HitTest(ByVal Token As Long, _
    ByVal X As Long, ByVal Y As Long) As Long
Public Sub OnMouse(ByVal Token As Long, ByVal Kind As MouseEventKind, _
    ByVal Region As Long, ByVal Button As Integer, _
    ByVal Buttons As Integer, ByVal Shift As Integer, _
    ByVal X As Long, ByVal Y As Long, ByVal Reason As MouseCancelReason)
```

`MouseRegister` returns a nonzero opaque token; failure returns zero without a partial registration. Reject duplicate windows and invalid/foreign-thread handles. Tokens must distinguish registrations even when Windows reuses a handle. `MouseUnregister` and `MouseShutdown` are idempotent and silent teardown operations. `MouseCancel` reports cancellation while keeping the registration available.

Coordinates are signed client pixels; captured movement may be negative. Convert at the control boundary for existing twip-based drawing. `HitTest` returns zero for no actionable region, one for a normal button, or a positive CPU index. The manager first checks window eligibility and actual pointer ownership; it must not accept an obscured region merely because its rectangle contains the point. Custom `Enabled` setters call `MouseSetEnabled`, because the existing properties do not set the native window's enabled state. Layout changes call `MouseRefresh` to reevaluate a stationary pointer.

Proposed event kinds are `Enter`, `Leave`, `Move`, `Down`, `Up`, `Click`, and `Cancel`. Enter/Leave refer to actionable regions: crossing between CPU regions produces Leave(old), Enter(new), then Move. Button identifies the changed button for Down/Up/Click; Buttons is the currently held mask. Other events use Button zero. Reason is nonzero only for Cancel. Add timed Hover later if a consumer needs it; it is distinct from immediate Enter.

For the first version, only left-button presses arm a click. A valid Down captures the window and records a positive region. Movement updates the current region without changing the origin. Up ends the press and releases owned capture before notifying consumers; Click follows Up only if release is over the same positive, eligible region and no cancellation or recipient mutation invalidated the gesture. Recheck registration after each callback. Cancellation clears the press and hover, emits Cancel then any necessary Leave, and never synthesizes Up or Click. Ordinary release must not become a spurious Cancel when Windows reports capture loss.

The registry holds ordinary object references, not reconstructed `ObjPtr` references. Owners explicitly unregister before unloading; native window destruction provides a second cleanup path. Do not rely solely on `Class_Terminate`, since a retained sink reference can prevent termination. Invalidate registrations before teardown, defer freeing records used by an active dispatch, and contain callback errors at the Windows boundary. A failed recipient should be detached and its error recorded for diagnosis.

**Classes and standard modules**

Classes and controls implement `IMouseSink` directly. For a standard `.bas` recipient, use a tiny application adapter class implementing the same interface and forwarding calls to, for example, `modConsumer.HandleMouse`. Register the adapter and retain its token in the module. This delivers notifications to modules, with explicit lifecycle management, without pretending that `.bas` modules can use `WithEvents`. Microsoft's classic-language references describe [WithEvents as class-only](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement) and prohibit [Event declarations in standard modules](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/event-statement); the exact interface will be compile-checked with VB6.

This proposal assumes “exportable module” can be a small source package with a module adapter when needed. If direct `AddressOf` registration or a single `.bas` file is essential, that changes the delivery design: passing a pointer is possible, but Basic cannot directly invoke it. A separate invocation mechanism would need validation. See [AddressOf](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/addressof-operator). The adapter is the recommended first version.

**Windows integration and first validation step**

Use native mouse messages and `TrackMouseEvent(TME_LEAVE)` for passive tracking; reserve capture for an active press. Leave tracking must be rearmed after reentry. See [TrackMouseEvent](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-trackmouseevent). Handle capture loss and cancellation explicitly; Windows sends `WM_CAPTURECHANGED` even after our own release, and it must not trigger immediate recapture. See [WM_CAPTURECHANGED](https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-capturechanged).

Prototype `SetWindowSubclass`/`RemoveWindowSubclass` with `DefSubclassProc` on registered control windows. This uses the Windows Common Controls DLL already relevant to this project, not an extra OCX. Verify API availability on supported deployment systems; Microsoft specifies Comctl32 5.8 or later and same-thread use. See [SetWindowSubclass](https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-setwindowsubclass) and [DefSubclassProc](https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-defsubclassproc).

The first prototype must establish message ordering around VB6's own default capture behavior and avoid duplicate notifications from old VB mouse handlers. It must also cover owner deactivation, hiding, destruction, and layout changes, using owner lifecycle forwarding where a child message is insufficient. `modHibernationAwareness.bas:17` currently replaces the main form's window procedure, and `modMain.bas:117` disables that hook in the IDE. Keep the mouse prototype on child windows and verify coexistence before introducing any owner hook. Do not claim IDE Stop/Reset safety for native callbacks; establish the IDE policy before migration.

Next, compile the unchanged project and build a small interaction harness that records notifications without launching stress workers. Verify: normal click; drag out/back; release outside; empty CPU space; crossing CPU regions; capture theft; disabling/hiding/unloading during a press or callback; deactivation; move/resize under a stationary cursor; multiple control instances; and repeated register/unregister. Test both class recipients and the module adapter, normal teardown, and callback failure. Confirm Button keyboard activation remains intact. Once that baseline and the subclass lifecycle work, migrate Button alone and compare behavior before proceeding.

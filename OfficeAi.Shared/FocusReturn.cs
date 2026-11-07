namespace OfficeAi.Shared
{
    public enum FocusReturnStep
    {
        Wait,    // keep watching
        Restore, // the app is back in front after being away: give the pane its focus again
        Stop,    // the pane page lost focus but the user never left the app (clicked the document): do nothing
    }

    // Decides what to do when the pane's page loses keyboard focus. Alt-tabbing to another program and back makes
    // Office hand focus to the document instead of the pane, so the cursor in the chat box is lost. A short timer
    // samples whether the pane's top-level window is in front; this decides from those samples (see PaneHostBase).
    public static class FocusReturn
    {
        public static FocusReturnStep Next(bool paneWindowIsInFront, ref bool wasAway)
        {
            if (!paneWindowIsInFront)
            {
                wasAway = true;
                return FocusReturnStep.Wait;
            }
            return wasAway ? FocusReturnStep.Restore : FocusReturnStep.Stop;
        }
    }
}

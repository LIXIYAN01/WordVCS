namespace WordVCS.AddIn
{
    /// <summary>
    /// Embedded Ribbon XML — defines the custom tab next to Home.
    /// </summary>
    internal static class RibbonXml
    {
        public const string Text = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabWordVCS"" label=""Paper VCS"" insertBeforeMso=""TabHome"">
        <group id=""gVer"" label=""Version"">
          <button id=""bShow"" label=""Show Panel"" size=""large""
                  onAction=""OnShowPanel"" imageMso=""ReviewShowAllComments""/>
          <separator id=""s1""/>
          <button id=""bCommit"" label=""Commit"" size=""normal""
                  onAction=""OnCommit"" imageMso=""FileSave""/>
          <button id=""bHistory"" label=""History"" size=""normal""
                  onAction=""OnHistory"" imageMso=""HistoryStore""/>
          <button id=""bDiff"" label=""Diff"" size=""normal""
                  onAction=""OnDiff"" imageMso=""CompareDocuments""/>
        </group>
        <group id=""gBranch"" label=""Branch / Tag"">
          <button id=""bBranch"" label=""Branches"" size=""large""
                  onAction=""OnBranch"" imageMso=""MergeCells""/>
          <button id=""bTag"" label=""Add Tag"" size=""normal""
                  onAction=""OnTag"" imageMso=""AddTag""/>
          <button id=""bFeedback"" label=""Import Feedback"" size=""normal""
                  onAction=""OnFeedback"" imageMso=""ReviewAcceptChange""/>
        </group>
        <group id=""gCfg"" label=""Settings"">
          <button id=""bSettings"" label=""Settings"" size=""large""
                  onAction=""OnSettings"" imageMso=""SourceControlSettings""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}

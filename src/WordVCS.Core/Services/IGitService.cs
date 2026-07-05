using System.Collections.Generic;
using System.Threading.Tasks;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// Git 版本控制服务接口
    /// </summary>
    public interface IGitService
    {
        bool IsRepository(string docPath);
        Task<bool> InitRepository(string docPath, string authorName, string authorEmail, string defaultBranch);
        Task<string> Commit(string docPath, string message, string authorName, string authorEmail);
        List<VersionRecord> GetHistory(string docPath, string branchName = null, int maxCount = 100);
        VersionRecord GetCommitDetail(string docPath, string commitSha);
        List<BranchInfo> ListBranches(string docPath);
        Task<bool> CreateBranch(string docPath, string branchName);
        Task<bool> SwitchBranch(string docPath, string branchName);
        Task<bool> CreateTag(string docPath, string tagName, string message);
        List<string> ListTags(string docPath);
        Task<bool> RestoreVersion(string docPath, string commitSha, string outputPath);
        bool HasUncommittedChanges(string docPath);
        string GetCurrentBranch(string docPath);
    }
}

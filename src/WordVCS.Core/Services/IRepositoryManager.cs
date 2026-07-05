using System.Collections.Generic;
using System.Threading.Tasks;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// WordVCS 仓库管理器接口，控制仓库层面的整体事务
    /// </summary>
    public interface IRepositoryManager
    {
        Task<bool> EnsureRepository(string docPath, string authorName,
            string authorEmail, string defaultBranch);

        Task<string> ExecuteCommit(string docPath, string message,
            string authorName, string authorEmail,
            List<CommentRecord> addressedComments);

        Task<string> TriggerAutoSave(string docPath, string authorName,
            string authorEmail);

        RepositoryStatus GetStatus(string docPath);
    }
}

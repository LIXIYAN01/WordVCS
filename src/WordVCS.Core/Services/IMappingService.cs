using System.Collections.Generic;
using System.Threading.Tasks;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// 批注 ↔ 提交映射关联的本地数据库管理服务接口
    /// </summary>
    public interface IMappingService
    {
        /// <summary>
        /// 初始化该论文仓库的专属数据库文件 wordvcs.db 并建表
        /// </summary>
        void InitializeDatabase(string docPath);

        /// <summary>
        /// 保存批注的最新状态及关联映射信息
        /// </summary>
        Task SaveMapping(string docPath, string commitSha,
            List<CommentRecord> addressedComments);

        /// <summary>
        /// 获取某个批注的映射信息
        /// </summary>
        CommentMapping GetMappingForComment(string docPath, string commentId);

        /// <summary>
        /// 获取某次提交中标记解决的所有批注列表
        /// </summary>
        List<CommentRecord> GetCommentsByCommit(string docPath, string commitSha);

        /// <summary>
        /// 批量导入外部批注（如导师反馈版的批注）
        /// </summary>
        Task ImportComments(string docPath, List<CommentRecord> comments,
            string firstSeenCommitSha);

        /// <summary>
        /// 从数据库中获取所有已登记的批注
        /// </summary>
        List<CommentRecord> GetPersistedComments(string docPath);

        /// <summary>
        /// 获取所有未处理的批注
        /// </summary>
        List<CommentRecord> GetUnresolvedComments(string docPath);

        /// <summary>
        /// 更新单条批注的状态
        /// </summary>
        Task UpdateCommentStatus(string docPath, string commentId,
            CommentStatus status, string commitSha = null);
    }
}

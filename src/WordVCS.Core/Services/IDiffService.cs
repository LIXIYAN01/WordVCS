using System.Threading.Tasks;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// 文档差异比较与高亮渲染服务接口
    /// </summary>
    public interface IDiffService
    {
        /// <summary>
        /// 从 .docx 文件中提取纯文本（保留段落结构）
        /// </summary>
        Task<string> ExtractTextFromDocx(string filePath);

        /// <summary>
        /// 比较两个纯文本字符串，生成行级/词级差异块
        /// </summary>
        Task<DiffResult> CompareText(string oldText, string newText);

        /// <summary>
        /// 比较两个本地 .docx 文档，计算整体差异
        /// </summary>
        Task<DiffResult> CompareDocx(string oldFilePath, string newFilePath);
    }
}

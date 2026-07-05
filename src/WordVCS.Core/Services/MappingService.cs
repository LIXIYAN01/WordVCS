using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Threading.Tasks;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// SQLite 数据库映射存储的具体实现
    /// </summary>
    public class MappingService : IMappingService
    {
        private static string GetDbPath(string docPath)
        {
            var dir = Path.GetDirectoryName(docPath);
            if (string.IsNullOrEmpty(dir)) return null;

            var wordvcsDir = Path.Combine(dir, ".wordvcs");
            if (!Directory.Exists(wordvcsDir))
                Directory.CreateDirectory(wordvcsDir);

            return Path.Combine(wordvcsDir, "wordvcs.db");
        }

        private static string GetConnectionString(string dbPath)
        {
            return $"Data Source={dbPath};Version=3;";
        }

        public void InitializeDatabase(string docPath)
        {
            var dbPath = GetDbPath(docPath);
            if (string.IsNullOrEmpty(dbPath)) return;

            var isNew = !File.Exists(dbPath);

            using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
            {
                conn.Open();

                if (isNew)
                {
                    const string createComments = @"
                        CREATE TABLE IF NOT EXISTS comments (
                            id TEXT PRIMARY KEY,
                            author TEXT NOT NULL,
                            content TEXT NOT NULL,
                            created_at TEXT NOT NULL,
                            range_start INTEGER DEFAULT 0,
                            range_end INTEGER DEFAULT 0,
                            selected_text TEXT DEFAULT '',
                            status TEXT DEFAULT 'Pending',
                            resolved_in_commit TEXT DEFAULT '',
                            first_seen_sha TEXT DEFAULT '',
                            last_updated TEXT NOT NULL
                        );";

                    const string createMappings = @"
                        CREATE TABLE IF NOT EXISTS comment_mappings (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            comment_id TEXT NOT NULL,
                            commit_sha TEXT NOT NULL,
                            action TEXT NOT NULL DEFAULT 'Addressed',
                            mapped_at TEXT NOT NULL,
                            FOREIGN KEY (comment_id) REFERENCES comments(id)
                        );";

                    const string createCommits = @"
                        CREATE TABLE IF NOT EXISTS commits_cache (
                            sha TEXT PRIMARY KEY,
                            message TEXT NOT NULL,
                            author TEXT NOT NULL,
                            committed_at TEXT NOT NULL
                        );";

                    using (var cmd = new SQLiteCommand(createComments, conn))
                        cmd.ExecuteNonQuery();
                    using (var cmd = new SQLiteCommand(createMappings, conn))
                        cmd.ExecuteNonQuery();
                    using (var cmd = new SQLiteCommand(createCommits, conn))
                        cmd.ExecuteNonQuery();

                    // Create indices
                    var indices = new[]
                    {
                        "CREATE INDEX IF NOT EXISTS idx_comments_status ON comments(status);",
                        "CREATE INDEX IF NOT EXISTS idx_mappings_comment ON comment_mappings(comment_id);",
                        "CREATE INDEX IF NOT EXISTS idx_mappings_commit ON comment_mappings(commit_sha);",
                    };
                    foreach (var idx in indices)
                    {
                        using (var cmd = new SQLiteCommand(idx, conn))
                            cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        public Task SaveMapping(string docPath, string commitSha,
            List<CommentRecord> addressedComments)
        {
            return Task.Run(() =>
            {
                var dbPath = GetDbPath(docPath);
                if (string.IsNullOrEmpty(dbPath)) return;

                using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    {
                        try
                        {
                            var now = DateTime.Now.ToString("o");

                            foreach (var comment in addressedComments)
                            {
                                // Upsert comment
                                const string upsertComment = @"
                                    INSERT OR REPLACE INTO comments
                                        (id, author, content, created_at,
                                         range_start, range_end, selected_text,
                                         status, resolved_in_commit,
                                         first_seen_sha, last_updated)
                                    VALUES
                                        (@id, @author, @content, @created_at,
                                         @range_start, @range_end, @selected_text,
                                         @status, @resolved_in_commit,
                                         @first_seen_sha, @last_updated);";

                                using (var cmd = new SQLiteCommand(upsertComment, conn))
                                {
                                    cmd.Parameters.AddWithValue("@id", comment.Id);
                                    cmd.Parameters.AddWithValue("@author", comment.Author ?? "");
                                    cmd.Parameters.AddWithValue("@content", comment.Content ?? "");
                                    cmd.Parameters.AddWithValue("@created_at", comment.CreatedAt.ToString("o"));
                                    cmd.Parameters.AddWithValue("@range_start", comment.RangeStart);
                                    cmd.Parameters.AddWithValue("@range_end", comment.RangeEnd);
                                    cmd.Parameters.AddWithValue("@selected_text", comment.SelectedText ?? "");
                                    cmd.Parameters.AddWithValue("@status", "Addressed");
                                    cmd.Parameters.AddWithValue("@resolved_in_commit", commitSha);
                                    cmd.Parameters.AddWithValue("@first_seen_sha", comment.FirstSeenSha ?? commitSha);
                                    cmd.Parameters.AddWithValue("@last_updated", now);
                                    cmd.ExecuteNonQuery();
                                }

                                // Insert mapping
                                const string insertMapping = @"
                                    INSERT INTO comment_mappings
                                        (comment_id, commit_sha, action, mapped_at)
                                    VALUES
                                        (@comment_id, @commit_sha, @action, @mapped_at);";

                                using (var cmd = new SQLiteCommand(insertMapping, conn))
                                {
                                    cmd.Parameters.AddWithValue("@comment_id", comment.Id);
                                    cmd.Parameters.AddWithValue("@commit_sha", commitSha);
                                    cmd.Parameters.AddWithValue("@action", "Addressed");
                                    cmd.Parameters.AddWithValue("@mapped_at", now);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            tx.Commit();
                        }
                        catch (Exception ex)
                        {
                            tx.Rollback();
                            System.Diagnostics.Debug.WriteLine(
                                $"[WordVCS] SaveMapping failed: {ex.Message}");
                        }
                    }
                }
            });
        }

        public CommentMapping GetMappingForComment(string docPath, string commentId)
        {
            var dbPath = GetDbPath(docPath);
            if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath)) return null;

            using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
            {
                conn.Open();
                const string sql = @"
                    SELECT * FROM comment_mappings
                    WHERE comment_id = @cid ORDER BY id DESC LIMIT 1;";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@cid", commentId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new CommentMapping
                            {
                                Id = Convert.ToInt32(reader["id"]),
                                CommentId = reader["comment_id"].ToString(),
                                CommitSha = reader["commit_sha"].ToString(),
                                Action = reader["action"].ToString(),
                                MappedAt = DateTime.Parse(reader["mapped_at"].ToString())
                            };
                        }
                    }
                }
            }
            return null;
        }

        public List<CommentRecord> GetCommentsByCommit(string docPath,
            string commitSha)
        {
            var list = new List<CommentRecord>();
            var dbPath = GetDbPath(docPath);
            if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath)) return list;

            using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
            {
                conn.Open();
                const string sql = @"
                    SELECT c.* FROM comments c
                    INNER JOIN comment_mappings m ON c.id = m.comment_id
                    WHERE m.commit_sha = @sha;";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@sha", commitSha);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            list.Add(ReadCommentRecord(reader));
                    }
                }
            }
            return list;
        }

        public Task ImportComments(string docPath, List<CommentRecord> comments,
            string firstSeenCommitSha)
        {
            return Task.Run(() =>
            {
                var dbPath = GetDbPath(docPath);
                if (string.IsNullOrEmpty(dbPath)) return;

                using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    {
                        try
                        {
                            var now = DateTime.Now.ToString("o");
                            foreach (var c in comments)
                            {
                                const string sql = @"
                                    INSERT OR IGNORE INTO comments
                                        (id, author, content, created_at,
                                         range_start, range_end, selected_text,
                                         status, first_seen_sha, last_updated)
                                    VALUES
                                        (@id, @author, @content, @created_at,
                                         @range_start, @range_end, @selected_text,
                                         'Pending', @first_seen_sha, @now);";

                                using (var cmd = new SQLiteCommand(sql, conn))
                                {
                                    cmd.Parameters.AddWithValue("@id", c.Id);
                                    cmd.Parameters.AddWithValue("@author", c.Author ?? "");
                                    cmd.Parameters.AddWithValue("@content", c.Content ?? "");
                                    cmd.Parameters.AddWithValue("@created_at", c.CreatedAt.ToString("o"));
                                    cmd.Parameters.AddWithValue("@range_start", c.RangeStart);
                                    cmd.Parameters.AddWithValue("@range_end", c.RangeEnd);
                                    cmd.Parameters.AddWithValue("@selected_text", c.SelectedText ?? "");
                                    cmd.Parameters.AddWithValue("@first_seen_sha", firstSeenCommitSha ?? "");
                                    cmd.Parameters.AddWithValue("@now", now);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            tx.Commit();
                        }
                        catch (Exception ex)
                        {
                            tx.Rollback();
                            System.Diagnostics.Debug.WriteLine(
                                $"[WordVCS] ImportComments failed: {ex.Message}");
                        }
                    }
                }
            });
        }

        public List<CommentRecord> GetPersistedComments(string docPath)
        {
            var list = new List<CommentRecord>();
            var dbPath = GetDbPath(docPath);
            if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath)) return list;

            using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
            {
                conn.Open();
                const string sql = "SELECT * FROM comments ORDER BY created_at DESC;";
                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                        list.Add(ReadCommentRecord(reader));
                }
            }
            return list;
        }

        public List<CommentRecord> GetUnresolvedComments(string docPath)
        {
            var list = new List<CommentRecord>();
            var dbPath = GetDbPath(docPath);
            if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath)) return list;

            using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
            {
                conn.Open();
                const string sql = @"
                    SELECT * FROM comments
                    WHERE status = 'Pending' OR status = 'Replied'
                    ORDER BY created_at ASC;";
                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                        list.Add(ReadCommentRecord(reader));
                }
            }
            return list;
        }

        public Task UpdateCommentStatus(string docPath, string commentId,
            CommentStatus status, string commitSha = null)
        {
            return Task.Run(() =>
            {
                var dbPath = GetDbPath(docPath);
                if (string.IsNullOrEmpty(dbPath)) return;

                using (var conn = new SQLiteConnection(GetConnectionString(dbPath)))
                {
                    conn.Open();
                    const string sql = @"
                        UPDATE comments SET status = @status,
                            resolved_in_commit = COALESCE(@sha, resolved_in_commit),
                            last_updated = @now
                        WHERE id = @id;";

                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@status", status.ToString());
                        cmd.Parameters.AddWithValue("@sha",
                            (object)commitSha ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@now", DateTime.Now.ToString("o"));
                        cmd.Parameters.AddWithValue("@id", commentId);
                        cmd.ExecuteNonQuery();
                    }
                }
            });
        }

        private static CommentRecord ReadCommentRecord(SQLiteDataReader reader)
        {
            return new CommentRecord
            {
                Id = reader["id"].ToString(),
                Author = reader["author"].ToString(),
                Content = reader["content"].ToString(),
                CreatedAt = SafeParseDate(reader["created_at"].ToString()),
                RangeStart = SafeParseInt(reader["range_start"].ToString()),
                RangeEnd = SafeParseInt(reader["range_end"].ToString()),
                SelectedText = reader["selected_text"]?.ToString() ?? "",
                Status = SafeParseEnum<CommentStatus>(reader["status"].ToString()),
                ResolvedInCommit = reader["resolved_in_commit"]?.ToString() ?? "",
                FirstSeenSha = reader["first_seen_sha"]?.ToString() ?? ""
            };
        }

        private static DateTime SafeParseDate(string s)
        {
            return DateTime.TryParse(s, out var d) ? d : DateTime.MinValue;
        }

        private static int SafeParseInt(string s)
        {
            return int.TryParse(s, out var i) ? i : 0;
        }

        private static T SafeParseEnum<T>(string s) where T : struct
        {
            return Enum.TryParse<T>(s, out var v) ? v : default;
        }
    }
}

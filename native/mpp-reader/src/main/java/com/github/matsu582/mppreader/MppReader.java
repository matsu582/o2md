package com.github.matsu582.mppreader;

import java.io.File;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import net.sf.mpxj.ProjectFile;
import net.sf.mpxj.ProjectProperties;
import net.sf.mpxj.Resource;
import net.sf.mpxj.ResourceAssignment;
import net.sf.mpxj.Task;
import net.sf.mpxj.Duration;
import net.sf.mpxj.reader.UniversalProjectReader;

/**
 * MS Projectファイルを読み込み、JSON形式で標準出力に出力するCLIツール。
 *
 * GraalVM Native Imageでネイティブバイナリにコンパイルし、
 * PythonからJRE不要で呼び出せるようにする。
 *
 * 使用方法: mpp-reader <ファイルパス>
 * 出力: JSON（タスク一覧・リソース一覧・プロジェクト情報）
 */
public class MppReader {

    private static final DateTimeFormatter DATE_FMT =
        DateTimeFormatter.ofPattern("yyyy/MM/dd");

    public static void main(String[] args) {
        if (args.length < 1) {
            System.err.println("使用方法: mpp-reader <ファイルパス>");
            System.exit(1);
        }

        String filePath = args[0];
        File file = new File(filePath);
        if (!file.exists()) {
            System.err.println("ファイルが見つかりません: " + filePath);
            System.exit(1);
        }

        try {
            ProjectFile project = new UniversalProjectReader().read(file);
            if (project == null) {
                System.err.println("ファイルの読み込みに失敗: " + filePath);
                System.exit(1);
            }

            PrintStream out = new PrintStream(System.out, true,
                StandardCharsets.UTF_8);
            writeJson(out, project);

        } catch (Exception e) {
            System.err.println("エラー: " + e.getMessage());
            System.exit(1);
        }
    }

    /**
     * プロジェクト情報をJSON形式で出力する。
     * 外部JSONライブラリに依存しないため手動で構築する。
     */
    private static void writeJson(PrintStream out, ProjectFile project) {
        out.println("{");

        // プロジェクト情報
        writeProjectInfo(out, project);
        out.println(",");

        // タスク一覧
        out.println("  \"tasks\": [");
        List<Task> topTasks = project.getChildTasks();
        writeTaskList(out, topTasks, true);
        out.println("  ],");

        // リソース一覧
        out.println("  \"resources\": [");
        writeResources(out, project);
        out.println("  ]");

        out.println("}");
    }

    private static void writeProjectInfo(PrintStream out,
                                         ProjectFile project) {
        ProjectProperties props = project.getProjectProperties();
        out.println("  \"project\": {");

        String title = safeStr(props.getProjectTitle());
        String author = safeStr(props.getAuthor());
        LocalDateTime start = props.getStartDate();
        LocalDateTime finish = props.getFinishDate();

        out.println("    \"title\": " + jsonStr(title) + ",");
        out.println("    \"author\": " + jsonStr(author) + ",");
        out.println("    \"start\": " + jsonStr(fmtDate(start)) + ",");
        out.println("    \"finish\": " + jsonStr(fmtDate(finish)));
        out.println("  }");
    }

    /**
     * タスクリストを再帰的にJSON配列として出力する。
     * 各タスクはフラットなオブジェクトで、outline_levelで階層を表現する。
     */
    private static void writeTaskList(PrintStream out, List<Task> tasks,
                                      boolean isFirst) {
        for (int i = 0; i < tasks.size(); i++) {
            Task task = tasks.get(i);
            String name = task.getName();
            if (name == null) continue;

            if (!isFirst) {
                out.println(",");
            }
            isFirst = false;

            int outline = 0;
            Number ol = task.getOutlineLevel();
            if (ol != null) outline = ol.intValue();

            LocalDateTime start = task.getStart();
            LocalDateTime finish = task.getFinish();
            Duration dur = task.getDuration();
            Number pct = task.getPercentageComplete();

            boolean hasSub = task.getChildTasks() != null
                && !task.getChildTasks().isEmpty();

            // 担当者の取得
            StringBuilder resBuf = new StringBuilder();
            List<ResourceAssignment> assignments =
                task.getResourceAssignments();
            for (int j = 0; j < assignments.size(); j++) {
                Resource res = assignments.get(j).getResource();
                if (res != null && res.getName() != null) {
                    if (resBuf.length() > 0) resBuf.append(", ");
                    resBuf.append(res.getName());
                }
            }

            out.print("    {");
            out.print("\"name\": " + jsonStr(name));
            out.print(", \"outline_level\": " + outline);
            out.print(", \"start\": " + jsonStr(fmtDate(start)));
            out.print(", \"finish\": " + jsonStr(fmtDate(finish)));
            out.print(", \"duration\": " + jsonStr(fmtDuration(dur)));
            out.print(", \"percent_complete\": " + jsonStr(fmtPercent(pct)));
            out.print(", \"resources\": " + jsonStr(
                resBuf.length() > 0 ? resBuf.toString() : "-"));
            out.print(", \"is_summary\": " + hasSub);
            out.print("}");

            // 子タスクの再帰処理
            if (hasSub) {
                out.println(",");
                writeTaskList(out, task.getChildTasks(), true);
            }
        }
    }

    private static void writeResources(PrintStream out,
                                       ProjectFile project) {
        List<Resource> resources = project.getResources();
        boolean first = true;
        for (Resource res : resources) {
            String name = res.getName();
            if (name == null) continue;

            if (!first) out.println(",");
            first = false;

            Number id = res.getID();
            int idVal = (id != null) ? id.intValue() : 0;

            out.print("    {\"id\": " + idVal
                + ", \"name\": " + jsonStr(name) + "}");
        }
        if (!first) out.println();
    }

    // --- ユーティリティ ---

    private static String fmtDate(LocalDateTime dt) {
        if (dt == null) return "-";
        return dt.format(DATE_FMT);
    }

    private static String fmtDuration(Duration dur) {
        if (dur == null) return "-";
        double val = dur.getDuration();
        String unit = dur.getUnits().toString();
        String label;
        switch (unit) {
            case "d": case "ed":   label = "日"; break;
            case "h": case "eh":   label = "時間"; break;
            case "w": case "ew":   label = "週"; break;
            case "mo": case "emo": label = "ヶ月"; break;
            case "m": case "em":   label = "分"; break;
            case "y": case "ey":   label = "年"; break;
            default:               label = unit; break;
        }
        if (val == (int) val) {
            return (int) val + label;
        }
        return String.format("%.1f%s", val, label);
    }

    private static String fmtPercent(Number pct) {
        if (pct == null) return "-";
        double val = pct.doubleValue();
        if (val == (int) val) {
            return (int) val + "%";
        }
        return String.format("%.1f%%", val);
    }

    private static String safeStr(String s) {
        return (s != null) ? s : "";
    }

    /**
     * 文字列をJSON文字列リテラルに変換する。
     * ダブルクォート・バックスラッシュ・制御文字をエスケープする。
     */
    private static String jsonStr(String s) {
        if (s == null) return "\"\"";
        StringBuilder sb = new StringBuilder("\"");
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '"':  sb.append("\\\""); break;
                case '\\': sb.append("\\\\"); break;
                case '\n': sb.append("\\n"); break;
                case '\r': sb.append("\\r"); break;
                case '\t': sb.append("\\t"); break;
                default:
                    if (c < 0x20) {
                        sb.append(String.format("\\u%04x", (int) c));
                    } else {
                        sb.append(c);
                    }
            }
        }
        sb.append("\"");
        return sb.toString();
    }
}

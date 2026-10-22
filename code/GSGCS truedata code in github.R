
###### 将服务器上保存的各基因筛选的rds结果整理为一个总excel ------
library(openxlsx)
cancerlist <- c("BRCA","LUAD","LUSC","STAD","OV","THYM")

type <- c("gsgc","coxcs")
out_path = paste0("C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/",
                  "260428结果/results_kegg-os/") # hallmark
# 创建列表存储所有数据
all_sheets <- list()
for (i in 1:length(cancerlist)) {
  for (j in 1:length(type)) {
    read_path <- paste0(out_path,"TCGA_",cancerlist[i],"_select_res_",type[j],".rds")    
    if (file.exists(read_path)) {
      data_read <- readRDS(read_path)
      # 提取 active_features（命名向量），转换为数据框，并保留基因名作为列
      theoretical_active <- data_read[["theoretical"]][["active_features"]]
      theoretical_df <- data.frame(
        type  = type[j],
        cancer= cancerlist[i], 
        Gene  = names(theoretical_active),
        order = 1:length(theoretical_active),
        Value = as.numeric(theoretical_active),
        stringsAsFactors = FALSE
      )
      # 添加其他元数据列（可选）
      
      theoretical_df$kappa_type <- data_read[["theoretical"]][["kappa_type"]]
      theoretical_df$kappa_name <- data_read[["theoretical"]][["kappa_name"]]
      theoretical_df$used_kappa <- data_read[["theoretical"]][["used_kappa"]]
      theoretical_df$threshold_value <- data_read[["theoretical"]][["threshold_value"]]
      
      # 提取 both 的 active_features
      both_active <- data_read[["both"]][["active_features"]]
      both_df <- data.frame(
        type  = type[j],
        cancer= cancerlist[i], 
        Gene = names(both_active),
        order = 1:length(both_active),
        Value = as.numeric(both_active),
        stringsAsFactors = FALSE
      )
      # 添加其他元数据列
      both_df$kappa_type <- data_read[["both"]][["kappa_type"]]
      both_df$kappa_name <- data_read[["both"]][["kappa_name"]]
      both_df$used_kappa <- data_read[["both"]][["used_kappa"]]
      both_df$threshold_value <- data_read[["both"]][["threshold_value"]]
      
      # 将数据添加到列表
      all_sheets[[paste0("k-os_",cancerlist[i],"-",type[j], "-theoretical")]] <- theoretical_df
      all_sheets[[paste0("k-os_",cancerlist[i],"-",type[j], "-Top k")]] <- both_df
    }
  }
}
# 一次性写入所有sheet
write_xlsx(all_sheets, paste0(out_path, "RNAexpression_truedata_kegg-os.xlsx"))


#### 将上述的excel按cancer拆开----

library(readxl)
# 你的目标目录（直接复制你给的路径）
target_dir <- "C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果"
# 列出该目录下所有 .xlsx / .xls 文件（不进入子目录）
excel_files <- list.files(
  path       = target_dir,
  pattern    = "\\.xlsx$|\\.xls$",  # 匹配 Excel 文件
  full.names = TRUE,                # 返回完整路径
  recursive  = TRUE                 # 不读取子目录
)
# 遍历每个 Excel 文件，输出文件名 + 所有 sheet 名
for (file in excel_files) {
  # 提取文件名（方便查看）
  file_name <- basename(file)
  # 获取该文件所有 sheet 名称
  sheet_names <- excel_sheets(file)
  # 打印结果
  cat("========================================\n")
  cat("文件：", file_name, "\n")
  cat("包含工作表：\n")
  cat(paste0("  - ", sheet_names, "\n"), "\n")
}
# ===================== 开始批量读取 + 自动加列 
all_data <- list()
for (file in excel_files) {
  # 文件名（用于解析 by 和 aim）
  file_name   <- basename(file)
  sheet_names <- excel_sheets(file)
  by_type     <- ifelse(grepl("kegg", file_name),"kegg", "hallmark")         # 解析 by：k=kegg, h=hallmark
  aim         <- ifelse(grepl("os", file_name), 
                        "os",ifelse(grepl("pfi", file_name), "pfi", "time")) # 解析 aim：os / pfi / time
  # 遍历每个 sheet
  for (sht in sheet_names) {
    # 读取数据，解析 sheet 名：cancer + 模型类型
    df         <- read_excel(file, sheet = sht)
    parts      <- strsplit(sht, "_")[[1]]
    info       <- strsplit(parts[2], "-")[[1]]
    cancer     <- info[1]
    model_type <- paste(info[-1], collapse = "-")
    # 自动加 2 列
    df$by         <- by_type
    df$aim        <- aim
    df$cancer     <- cancer
    df$model_type <- model_type
    # 存入列表
    key <- paste(cancer, by_type, aim, model_type, sep = "_") # paste0(cancer, "_", by_type, "_", aim, "_", model_type)
    all_data[[key]] <- df
  }
}
# ===================== 按癌症输出 5 个文件 
cancers     <- c("STAD", "OV", "BRCA", "LUAD", "THYM")
model_order <- c("gsgc-theoretical", "coxcs-theoretical", "gsgc-Top k", "coxcs-Top k")
by_aim_list <- c("hallmark_os", "hallmark_pfi", "hallmark_time", 
                 "kegg_os", "kegg_pfi", "kegg_time")
for (cancer in cancers) {
  sheets_list <- list()
  total_df    <- NULL
  # 匹配当前癌症的所有数据
  keys <- names(all_data)[grepl(paste0("^", cancer, "_"), names(all_data))]
  for (key in keys) {
    sheets_list[[key]] <- all_data[[key]]
    total_df <- bind_rows(total_df, all_data[[key]])
    cat(cancer, "-", key, ":", nrow(all_data[[key]]), "行\n")
  }
  # 加入总表
  if (!is.null(total_df)) {
    sheets_list[["total"]] <- total_df
  }
  # 输出到桌面（方便查看）
  out_path <- paste0("C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/",
                     "260428结果/",
                     "RNAexpression_truedata_", cancer, "-new.xlsx")
  write_xlsx(sheets_list, path = out_path)
  cat("已生成：", out_path, "\n")
}

##### 文献中报道的明星基因检出率
library(openxlsx)
library(dplyr)
library(tidyr)

# 定义参数
cancerlist  <- c("BRCA","LUAD","STAD","OV","THYM")
sheet_types <- c("total")  # c("theoretical", "Top k")  # 两种sheet类型
out_path    <- paste0("C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/",
                      "260428结果/")
# 读取文献基因数据（只需要读取一次）
genereferdataset <- read.xlsx(paste0(out_path, "文献报道癌种相关基因汇总.xlsx"),sheet = "文献报道gene")  # sheet = TCGA实验验证gene
# 循环遍历每个癌症
for (cancer in cancerlist) {
  # 循环遍历每种sheet类型
  for(sheet_type in sheet_types) {
    cat("\n========================================\n")
    cat(paste0("正在处理: ", cancer, " - ", sheet_type, "\n"))
    cat("========================================\n")
    # 构建文件路径和sheet名称
    file_name  <- paste0("RNAexpression_truedata_", cancer, "-new.xlsx")
    sheet_name <- paste0(sheet_type) # paste0(cancer, "-", sheet_type)
    file_path  <- paste0(out_path,file_name) # paste0(out_path, cancer, "/", file_name)
    # 检查文件是否存在
    if(!file.exists(file_path)) {
      cat(paste0("警告: 文件不存在 - ", file_path, "\n"))
      next
    }
    # 读取数据
    cancerdataset <- tryCatch({
      read.xlsx(file_path, sheet = sheet_name)
    }, error = function(e) {
      cat(paste0("错误: 无法读取sheet - ", sheet_name, "\n"))
      cat(paste0("错误信息: ", e$message, "\n"))
      return(NULL)
    })
    if(is.null(cancerdataset)) next
    # 确保必要的列存在
    required_cols <- c("type", "by", "aim", "cancer", "Gene", "Value")
    if(!all(required_cols %in% colnames(cancerdataset))) {
      cat(paste0("错误: 数据框缺少必要的列\n"))
      cat(paste0("现有列: ", paste(colnames(cancerdataset), collapse=", "), "\n"))
      next
    }
    # 方法：先拆分数据，再逐个分组处理
    result_list <- list()
    # 获取所有分组
    groups <- cancerdataset %>%
      group_by(type, by, aim, cancer,kappa_type) %>%
      group_keys()
    # 如果没有分组，跳过
    if(nrow(groups) == 0) {
      cat("警告: 没有找到分组数据\n")
      next
    }
    # 遍历每个分组
    for(i in 1:nrow(groups)) {
      # 获取当前分组的条件
      current_type <- groups$type[i]
      current_by <- groups$by[i]
      current_aim <- groups$aim[i]
      current_kappa_type <- groups$kappa_type[i]
      current_cancer <- groups$cancer[i]
      # 筛选当前分组的数据
      group_data <- cancerdataset %>%
        filter(type == current_type,
               by == current_by,
               aim == current_aim,
               kappa_type == current_kappa_type,
               cancer == current_cancer) %>%
        mutate(Gene = toupper(Gene))
      # 获取该癌症的文献基因
      ref_genes <- genereferdataset$gene[genereferdataset$cancer == current_cancer & genereferdataset$`by-aim` == paste0(current_by,"-",current_aim)]
      # 如果没有文献基因，跳过
      if(length(ref_genes) == 0) {
        cat(paste0("警告: ", current_cancer, " 没有文献基因\n"))
        next
      }
      # 检查每个文献基因
      detection_result <- data.frame()
      for(gene in ref_genes) {
        gene_upper <- toupper(gene)
        match_row <- group_data[group_data$Gene == gene_upper, ]
        if(nrow(match_row) > 0) {
          detection_result <- rbind(detection_result, data.frame(
            detected = TRUE,
            rank = match_row$Value[1]
          ))
        } else {
          detection_result <- rbind(detection_result, data.frame(
            detected = FALSE,
            rank = NA
          ))
        }
      }
      # 存储结果
      result_list[[i]] <- data.frame(
        type       = current_type,
        by         = current_by,
        aim        = current_aim,
        cancer     = current_cancer,
        kappa_type = current_kappa_type,
        sheet_type = sheet_type,  # 添加sheet类型标识
        literature_gene = ref_genes,
        detected = detection_result$detected,
        rank = detection_result$rank,
        stringsAsFactors = FALSE
      )
    }
    # 合并所有结果
    if(length(result_list) == 0) {
      cat("警告: 没有生成任何结果\n")
      next
    }
    final_result <- do.call(rbind, result_list)
    # 计算每个分组的检出率和排名统计
    summary_stats <- final_result %>%
      group_by(type, by, aim, cancer, kappa_type, sheet_type) %>%
      summarise(
        total_genes = n(),
        detected_genes = sum(detected),
        detection_rate = detected_genes / total_genes * 100,
        
        # 仅对检出的基因计算排名统计
        median_rank = median(rank[detected], na.rm = TRUE),
        mean_rank   = mean(rank[detected], na.rm = TRUE),
        min_rank    = min(rank[detected], na.rm = TRUE),
        max_rank    = max(rank[detected], na.rm = TRUE),
        
        # 排名在不同区间的分布
        rank_top50  = sum(rank <= 50 & detected, na.rm = TRUE),
        rank_top100 = sum(rank <= 100 & detected, na.rm = TRUE),
        rank_top500 = sum(rank <= 500 & detected, na.rm = TRUE),
        
        .groups = 'drop'
      )
    
    # 创建Excel工作簿
    wb <- createWorkbook()
    
    # 1. 详细检测结果
    addWorksheet(wb, "详细检测结果")
    writeData(wb, "详细检测结果", final_result)
    
    # 2. 汇总统计
    addWorksheet(wb, "汇总统计")
    writeData(wb, "汇总统计", summary_stats)
    
    # 3. 每个癌症的检出基因及其排名（宽格式）
    detected_genes_wide <- final_result %>%
      filter(detected == TRUE) %>%
      dplyr::select(-detected) %>%
      pivot_wider(
        names_from = literature_gene,
        values_from = rank,
        values_fill = list(rank = NA)
      )
    
    if(nrow(detected_genes_wide) > 0) {
      addWorksheet(wb, "检出基因排名")
      writeData(wb, "检出基因排名", detected_genes_wide)
    }
    
    # 4. 按癌症查看汇总结果
    cancer_summary <- final_result %>%
      group_by(type, by, aim, cancer,kappa_type,  sheet_type) %>%
      summarise(
        总文献基因数 = n_distinct(literature_gene),
        平均检出率 = mean(detected) * 100,
        中位排名 = median(rank[detected], na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      arrange(desc(平均检出率))
    
    addWorksheet(wb, "癌症汇总")
    writeData(wb, "癌症汇总", cancer_summary)
    
    # 5. 查看未检出的基因
    missing_genes <- final_result %>%
      filter(detected == FALSE) %>%
      dplyr::select(type, by, aim, cancer, kappa_type,sheet_type, literature_gene) %>%
      distinct() %>%
      arrange(cancer, type, by, aim,kappa_type)
    
    if(nrow(missing_genes) > 0) {
      addWorksheet(wb, "未检出基因")
      writeData(wb, "未检出基因", missing_genes)
    }
    
    # 保存文件
    output_file <- paste0("基因检出率分析结果-", cancer, "-", sheet_type, ".xlsx")
    saveWorkbook(wb, output_file, overwrite = TRUE)
    
    # 输出完成提示
    cat("\n分析完成！结果已保存到:", output_file, "\n")
    cat(paste0("共分析了 ", nrow(groups), " 个分组\n"))
    cat(paste0("文献基因总数: ", length(unique(final_result$literature_gene)), "\n"))
    cat(paste0("总检测记录数: ", nrow(final_result), "\n"))
    cat(paste0("总检出次数: ", sum(final_result$detected), "\n"))
    cat(paste0("总未检出次数: ", sum(!final_result$detected), "\n"))
    cat(paste0("总体检出率: ", round(sum(final_result$detected)/nrow(final_result)*100, 2), "%\n"))
    
    # 额外输出一些关键信息
    cat("\n=== 各分组检出率汇总 ===\n")
    print(summary_stats %>% 
            dplyr::select(type, by, aim, cancer, kappa_type,detection_rate) %>%
            arrange(desc(detection_rate)))
    
    cat("\n========================================\n")
  }
}

# 读取文献基因数据（只需要读取一次）
genereferdataset <- read.xlsx(paste0(out_path, "文献报道癌种相关基因汇总.xlsx"),sheet = "TCGA实验验证gene")  # sheet = TCGA实验验证gene
# 循环遍历每个癌症
for (cancer in cancerlist) {
  # 循环遍历每种sheet类型
  for(sheet_type in sheet_types) {
    cat("\n========================================\n")
    cat(paste0("正在处理: ", cancer, " - ", sheet_type, "\n"))
    cat("========================================\n")
    # 构建文件路径和sheet名称
    file_name  <- paste0("RNAexpression_truedata_", cancer, "-new.xlsx")
    sheet_name <- paste0(sheet_type) # paste0(cancer, "-", sheet_type)
    file_path  <- paste0(out_path,file_name) # paste0(out_path, cancer, "/", file_name)
    # 检查文件是否存在
    if(!file.exists(file_path)) {
      cat(paste0("警告: 文件不存在 - ", file_path, "\n"))
      next
    }
    # 读取数据
    cancerdataset <- tryCatch({
      read.xlsx(file_path, sheet = sheet_name)
    }, error = function(e) {
      cat(paste0("错误: 无法读取sheet - ", sheet_name, "\n"))
      cat(paste0("错误信息: ", e$message, "\n"))
      return(NULL)
    })
    if(is.null(cancerdataset)) next
    # 确保必要的列存在
    required_cols <- c("type", "by", "aim", "cancer", "Gene", "Value")
    if(!all(required_cols %in% colnames(cancerdataset))) {
      cat(paste0("错误: 数据框缺少必要的列\n"))
      cat(paste0("现有列: ", paste(colnames(cancerdataset), collapse=", "), "\n"))
      next
    }
    # 方法：先拆分数据，再逐个分组处理
    result_list <- list()
    # 获取所有分组
    groups <- cancerdataset %>%
      group_by(type, by, aim, cancer,kappa_type) %>%
      group_keys()
    # 如果没有分组，跳过
    if(nrow(groups) == 0) {
      cat("警告: 没有找到分组数据\n")
      next
    }
    # 遍历每个分组
    for(i in 1:nrow(groups)) {
      # 获取当前分组的条件
      current_type <- groups$type[i]
      current_by <- groups$by[i]
      current_aim <- groups$aim[i]
      current_kappa_type <- groups$kappa_type[i]
      current_cancer <- groups$cancer[i]
      # 筛选当前分组的数据
      group_data <- cancerdataset %>%
        filter(type == current_type,
               by == current_by,
               aim == current_aim,
               kappa_type == current_kappa_type,
               cancer == current_cancer) %>%
        mutate(Gene = toupper(Gene))
      # 获取该癌症的文献基因
      ref_genes <- genereferdataset$gene[genereferdataset$cancer == current_cancer & genereferdataset$`by-aim` == paste0(current_by,"-",current_aim)]
      # 如果没有文献基因，跳过
      if(length(ref_genes) == 0) {
        cat(paste0("警告: ", current_cancer, " 没有文献基因\n"))
        next
      }
      # 检查每个文献基因
      detection_result <- data.frame()
      for(gene in ref_genes) {
        gene_upper <- toupper(gene)
        match_row <- group_data[group_data$Gene == gene_upper, ]
        if(nrow(match_row) > 0) {
          detection_result <- rbind(detection_result, data.frame(
            detected = TRUE,
            rank = match_row$Value[1]
          ))
        } else {
          detection_result <- rbind(detection_result, data.frame(
            detected = FALSE,
            rank = NA
          ))
        }
      }
      # 存储结果
      result_list[[i]] <- data.frame(
        type       = current_type,
        by         = current_by,
        aim        = current_aim,
        cancer     = current_cancer,
        kappa_type = current_kappa_type,
        sheet_type = sheet_type,  # 添加sheet类型标识
        literature_gene = ref_genes,
        detected = detection_result$detected,
        rank = detection_result$rank,
        stringsAsFactors = FALSE
      )
    }
    # 合并所有结果
    if(length(result_list) == 0) {
      cat("警告: 没有生成任何结果\n")
      next
    }
    final_result <- do.call(rbind, result_list)
    # 计算每个分组的检出率和排名统计
    summary_stats <- final_result %>%
      group_by(type, by, aim, cancer, kappa_type, sheet_type) %>%
      summarise(
        total_genes = n(),
        detected_genes = sum(detected),
        detection_rate = detected_genes / total_genes * 100,
        
        # 仅对检出的基因计算排名统计
        median_rank = median(rank[detected], na.rm = TRUE),
        mean_rank   = mean(rank[detected], na.rm = TRUE),
        min_rank    = min(rank[detected], na.rm = TRUE),
        max_rank    = max(rank[detected], na.rm = TRUE),
        
        # 排名在不同区间的分布
        rank_top50  = sum(rank <= 50 & detected, na.rm = TRUE),
        rank_top100 = sum(rank <= 100 & detected, na.rm = TRUE),
        rank_top500 = sum(rank <= 500 & detected, na.rm = TRUE),
        
        .groups = 'drop'
      )
    
    # 创建Excel工作簿
    wb <- createWorkbook()
    
    # 1. 详细检测结果
    addWorksheet(wb, "详细检测结果")
    writeData(wb, "详细检测结果", final_result)
    
    # 2. 汇总统计
    addWorksheet(wb, "汇总统计")
    writeData(wb, "汇总统计", summary_stats)
    
    # 3. 每个癌症的检出基因及其排名（宽格式）
    detected_genes_wide <- final_result %>%
      filter(detected == TRUE) %>%
      dplyr::select(-detected) %>%
      pivot_wider(
        names_from = literature_gene,
        values_from = rank,
        values_fill = list(rank = NA)
      )
    
    if(nrow(detected_genes_wide) > 0) {
      addWorksheet(wb, "检出基因排名")
      writeData(wb, "检出基因排名", detected_genes_wide)
    }
    
    # 4. 按癌症查看汇总结果
    cancer_summary <- final_result %>%
      group_by(type, by, aim, cancer,kappa_type,  sheet_type) %>%
      summarise(
        总文献基因数 = n_distinct(literature_gene),
        平均检出率 = mean(detected) * 100,
        中位排名 = median(rank[detected], na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      arrange(desc(平均检出率))
    
    addWorksheet(wb, "癌症汇总")
    writeData(wb, "癌症汇总", cancer_summary)
    
    # 5. 查看未检出的基因
    missing_genes <- final_result %>%
      filter(detected == FALSE) %>%
      dplyr::select(type, by, aim, cancer, kappa_type,sheet_type, literature_gene) %>%
      distinct() %>%
      arrange(cancer, type, by, aim,kappa_type)
    
    if(nrow(missing_genes) > 0) {
      addWorksheet(wb, "未检出基因")
      writeData(wb, "未检出基因", missing_genes)
    }
    
    # 保存文件
    output_file <- paste0("-同数据库文章-基因检出率分析结果-", cancer, "-", sheet_type, ".xlsx")
    saveWorkbook(wb, output_file, overwrite = TRUE)
    
    # 输出完成提示
    cat("\n分析完成！结果已保存到:", output_file, "\n")
    cat(paste0("共分析了 ", nrow(groups), " 个分组\n"))
    cat(paste0("文献基因总数: ", length(unique(final_result$literature_gene)), "\n"))
    cat(paste0("总检测记录数: ", nrow(final_result), "\n"))
    cat(paste0("总检出次数: ", sum(final_result$detected), "\n"))
    cat(paste0("总未检出次数: ", sum(!final_result$detected), "\n"))
    cat(paste0("总体检出率: ", round(sum(final_result$detected)/nrow(final_result)*100, 2), "%\n"))
    
    # 额外输出一些关键信息
    cat("\n=== 各分组检出率汇总 ===\n")
    print(summary_stats %>% 
            dplyr::select(type, by, aim, cancer, kappa_type,detection_rate) %>%
            arrange(desc(detection_rate)))
    
    cat("\n========================================\n")
  }
}

cat("\n\n========== 所有分析完成！ ==========\n")


#### 补充通路富集分析


#### 补充通路富集分析 -----
# 加载必要的包
# 1. 安装并加载必需包
library(clusterProfiler)
library(enrichplot)
library(org.Hs.eg.db)
library(dplyr)
library(readxl)
library(tidyr)
library(dplyr)
library(ggplot2)
library(patchwork)

# 读取数据（关键：替换为你的文件路径）
# 癌症列表
cancer_list <- c("BRCA", "OV", "STAD", "THYM", "LUAD")
# 初始化空列表，用于存储每个癌症合并后的数据
all_data_list <- list()


for (i in 1:length(cancer_list)){
  input_file  <- paste0("C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/",
                        # cancer_list[i],"/",
                        "RNAexpression_truedata_",cancer_list[i],"-new.xlsx") # 你的数据路径
  # # 合并当前癌症的两个sheet
  cancer_combined <- read_excel(input_file, sheet = "total")  
  # 存储到列表中
  all_data_list[[i]] <- cancer_combined
}
# 合并所有癌症的数据
all_data <- bind_rows(all_data_list)

# 查看结果
str(all_data)
table(all_data$cancer)  # 查看各癌症样本量
table(all_data$kappa_type)  # 查看理论阈值和Top-k的样本量

# 3. 数据清洗（必须：确保列名匹配你的数据）
# 假设数据列名：gene（基因名，如TP53）、type（方法：GSGC/CoxCS）、by（通路：KEGG/Hallmark）、aim（结局：OS/PFI/time）
# 若列名不同，需修改下方"gene"、"type"、"by"、"aim"为你的实际列名
clean_data <- all_data %>%
  filter(!is.na(Gene)) %>%  # 去除无基因名的行
  filter(!is.na(type)) %>%  # 去除无方法的行
  filter(!is.na(by)) %>%    # 去除无通路的行
  filter(!is.na(aim)) %>%   # 去除无结局的行
  mutate(
    type = factor(type, levels = c("gsgc", "coxcs")),       # 固定方法顺序
    by   = factor(by,   levels = c("kegg", "hallmark")),    # 固定通路顺序
    aim  = factor(aim,  levels = c("os", "pfi", "time")),   # 固定结局顺序
    kappa_type= factor(kappa_type,  levels = c("theoretical", "Top‑k"))
  )
str(clean_data)

# （1）富集分析     -------------------------------------------------- ----
# 自动下载 Hallmark
if(!file.exists("C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260325结果/h.all.v2023.2.Hs.symbols.gmt")){
  download.file("https://data.broadinstitute.org/gsea-msigdb/msigdb/release/2023.2.Hs/h.all.v2023.2.Hs.symbols.gmt",
                "h.all.v2023.2.Hs.symbols.gmt")
}
hall <- read.gmt("C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260325结果/h.all.v2023.2.Hs.symbols.gmt")

# 富集函数（必须加ID转换）
run_enrich <- function(genes, db){
  genes <- unique(genes[!is.na(genes)])
  if(length(genes) < 5) return(NULL)
  
  if(db == "hallmark") {
    er <- enricher(
      gene          = genes,
      TERM2GENE     = hall,
      pAdjustMethod = "BH",
      pvalueCutoff  = 0.05,
      qvalueCutoff  = 0.1
    )
  } else {
    # KEGG 必须转 ENTREZ ID，否则不出结果！
    ids <- bitr(genes, fromType = "SYMBOL", toType = "ENTREZID", OrgDb = org.Hs.eg.db)
    er <- enrichKEGG(
      gene          = ids$ENTREZID,
      organism      = "hsa",
      pvalueCutoff  = 0.05,
      pAdjustMethod = "BH"    # 补充KEGG的BH校正
    )
    er <- setReadable(er, OrgDb = org.Hs.eg.db, keyType = "ENTREZID")
  }
  return(er)
}

run_enrich_id <- function(genes, db){
  genes <- unique(genes[!is.na(genes)])
  if(length(genes) < 5) return(NULL)
  
  if(db != "hallmark") {
    # KEGG 必须转 ENTREZ ID，否则不出结果！
    ids <- bitr(genes, fromType = "SYMBOL", toType = "ENTREZID", OrgDb = org.Hs.eg.db)
    er <- enrichKEGG(
      gene          = ids$ENTREZID,
      organism      = "hsa",
      pvalueCutoff  = 0.05,
      pAdjustMethod = "BH"    # 补充KEGG的BH校正
    )
    er <- setReadable(er, OrgDb = org.Hs.eg.db, keyType = "ENTREZID")
  }
  return(ids)
}


library(ggplot2)
library(dplyr)
library(clusterProfiler)
library(org.Hs.eg.db)
library(ggnewscale)  # 用于多图层

# 1. 数据准备和富集分析
cancer_list <- unique(clean_data$cancer)
methods     <- c("gsgc", "coxcs")
aims        <- c("os", "pfi", "time")
kappa_types <- c("theoretical", "Top‑k")

# 检查基因是否有重复
all_data_list[[1]] %>% 
  group_by(cancer, type, kappa_type, aim, by, model_type) %>% 
  count(Gene)


# 存储所有富集结果
enrich_results <- list()
for(cancer in cancer_list) {
  for(method in methods) {
    for(aim in aims) {
      for(kappa in kappa_types) {
        
        # 筛选Top100基因
        genes <- clean_data %>%
          # filter(order <= 100) %>%
          filter(
            cancer     == !!cancer,
            type       == !!method,
            aim        == !!aim,
            kappa_type == !!kappa
          ) %>%
          pull(Gene) %>%
          unique()
        
        if(length(genes) < 3) next
        
        # 1. Hallmark富集
        enrich_hall <- tryCatch({ run_enrich(genes, "hallmark") }, error = function(e) NULL)
        
        if(!is.null(enrich_hall)) {
          df_hall <- as.data.frame(enrich_hall)  # S4转普通数据框
          if(nrow(df_hall) == 0) next
          
          # 添加分组标签
          df_hall <- df_hall %>%
            mutate(
              by         = "hallmark",
              cancer     = cancer,
              type       = method,
              aim        = aim,
              kappa_type = kappa
            )
          enrich_results[[length(enrich_results)+1]] <- df_hall
        }
        
        # 2. KEGG富集
        enrich_kegg <- tryCatch({ run_enrich(genes, "kegg") }, error = function(e) NULL)
        
        if(!is.null(enrich_kegg)) {
          df_kegg <- as.data.frame(enrich_kegg)
          if(nrow(df_kegg) == 0) next
          
          df_kegg <- df_kegg %>%
            mutate(
              by         = "kegg",
              cancer     = cancer,
              type       = method,
              aim        = aim,
              kappa_type = kappa
            )
          enrich_results[[length(enrich_results)+1]] <- df_kegg
        }
      }
    }
  }
}

# 合并所有富集结果
all_enrich <- bind_rows(enrich_results)
all_enrich <- all_enrich %>%
  separate(GeneRatio, into = c("gene_in_set", "pathway_total"), sep = "/", convert = TRUE) %>%
  mutate(
    # 计算 GeneRatio_value（基因集中该通路的比例）
    GeneRatio_value = gene_in_set / pathway_total,
    LOGP            = -log10(p.adjust),
    # 解析 BgRatio
    bg_gene_in_set  = as.numeric(sapply(strsplit(BgRatio, "/"), `[`, 1)),
    bg_total        = as.numeric(sapply(strsplit(BgRatio, "/"), `[`, 2)),
    BgRatio_value   = bg_gene_in_set / bg_total,              # 背景比率（标准）
    fold_enrichment = GeneRatio_value / BgRatio_value,        # 标准富集倍数（Fold Enrichment
    kappa_type = case_when(kappa_type == "theoretical" ~ "Theoretical",TRUE ~ kappa_type)    
  ) %>%
  dplyr::select(
    cancer, type, aim, kappa_type, by,
    Description, ID,
    BgRatio,  # 保留 BgRatio
    P.values = pvalue,
    Padj     = p.adjust,
    LOGP,
    Count, geneID,
    gene_in_set, pathway_total, GeneRatio_value,
    bg_gene_in_set, bg_total, BgRatio_value,
    fold = fold_enrichment   # 重命名为 fold（方便绘图）
  )

# 检查 fold 是否合理（通常 >1 表示富集）
summary(all_enrich$fold)

# 查看结构
str(all_enrich)
head(all_enrich)

# （2）富集结果整理 -------------------------------------------------- ----
library(ggplot2)
library(dplyr)
library(forcats)
# 选择要画的数据：取显著的通路（可自己改阈值）
pdata <- all_enrich %>%
  filter(Padj < 0.05) %>%          # 显著富集
  group_by(cancer, type, aim, by, kappa_type, Description) %>%
  slice_head(n=1) %>%              # 去重
  ungroup()

# # 一般做富集的气泡图用gene ratio表示气泡大小，颜色表示P。
# # 本次使用logP表示气泡大小，fold表示颜色。
# pdata <- pdata %>% mutate(fold = -log(GeneRatio_value))  # 按你的公式计算fold

# 每个分组只展示 TOP 10 通路（避免图太挤）
top_num <- 10
pdata <- pdata %>%
  group_by(cancer, type, aim, by, kappa_type) %>%
  slice_min(order_by = Padj, n = top_num) %>%
  ungroup()

library(ggplot2)
library(dplyr)
library(patchwork)
library(ggnewscale)
library(wesanderson)

# 获得统一绘图数据
for(cancer1 in cancer_list) {
  
  library(ggplot2)
  library(dplyr)
  library(patchwork)
  library(ggnewscale)
  library(wesanderson)
  
  plotdata = pdata %>%
    filter(cancer==cancer1) %>%# 去掉 type 过滤 by=="kegg", 
    arrange(by, Description) %>%
    mutate(Description = factor(Description, levels = unique(Description)))
  
  # 1. 左侧色块（零边距）
  cluster <- plotdata %>%
    dplyr::select(Description, by) %>%
    distinct() %>%
    mutate(p = "") %>%
    ggplot(aes(p, Description, fill = by)) +
    geom_tile(color = NA) +  # 去掉边框，无缝隙  geom_tile() +
    scale_fill_manual(values = c(
      "hallmark" = "#6699CC",   # 莫兰迪紫灰"#A997B5"
      "kegg"     = "#CC3333"    # 莫兰迪蓝灰"#92A2B3"
      # "hallmark"="#E64B35",
      # "kegg"="#4DBBD5"
    )) +
    # 关键：Y轴和 p1 完全一样,与主图一致，无多余留白
    scale_y_discrete(position = "right",expand = c(0.02, 0)) + 
    # 关键：x轴是离散型，不能用 continuous！
    scale_x_discrete(expand = c(0, 0)) +
    theme_void() +
    theme(
      legend.position = "none",
      plot.margin = unit(c(0,0,0,0), "pt"),  # 🔥 零边距
      axis.text.y = element_blank(),         # 隐藏色块图的Y轴文本，避免重复
      axis.text.x = element_blank(),
      strip.text = element_blank(),          # 隐藏分面标题
      panel.spacing = unit(0, "cm"),         # 无内部间距
    )
  # 2. 主图（零边距）
  p1 <- plotdata %>%
    ggplot(aes(aim, Description)) +
    geom_point(aes(color = fold, size = LOGP, shape = type), alpha=0.85, stroke=0.8) +
    facet_wrap(~kappa_type, ncol = 2) +
    scale_color_gradientn(
      colours = RColorBrewer::brewer.pal(11, "Spectral"), # RColorBrewer::brewer.pal(9, "RdBu"), # RColorBrewer::brewer.pal(n=5, name = 'Set3') , #
      name = "Fold"
    ) +
    scale_size_continuous(name="LOGP", range=c(1.5,7)) +
    scale_shape_manual(values = c("gsgc" = 16, "coxcs" = 17), name = "Method") + # values = c("gsgc" = 1, "coxcs" = 3), name = "Method") +
    # 额外添加Pathway图例
    new_scale_fill() +
    geom_point(
      data = data.frame(by=c("hallmark","kegg")),
      aes(fill=by), x=1, y=1, size=0, show.legend=T, inherit.aes=F
    ) +
    scale_fill_manual(values = c(
      "hallmark" = "#6699CC",   # 莫兰迪紫灰"#A997B5"
      "kegg"     = "#CC3333"    # 莫兰迪蓝灰"#92A2B3"
    ), name="Pathway") +
    
    guides(
      shape = guide_legend(order=1),
      size  = guide_legend(order=2),
      color = guide_colorbar(order=3),
      fill  = guide_legend(
        order=4,
        byrow = F,
        # key.height   = unit(8,"pt"),
        override.aes = list(shape=22, size=8, alpha=1, color=NA)
      )
    ) +
    scale_y_discrete(
      position = "right",                                  # 主 y 轴在右侧（原来就在右侧）
      sec.axis = dup_axis(name = NULL, labels = waiver()), # 复制左侧 y 轴
      expand = c(0.02, 0)
    ) +
    labs(x=NULL, y=NULL) +
    theme_bw() +
    theme(
      axis.text.x = element_text(angle=0, hjust=1, size=25), # X轴刻度标签
      strip.text  = element_text(size=30, face="bold"),      # 分面标题
      # 放大右侧通路描述字体
      # axis.text.y = element_text(size=15),    # 通路名称的字体大小，你可以改成11/12/14
      axis.text.y.left  = element_text(size=25, lineheight = 0.9),  
      axis.text.y.right = element_text(size=25, lineheight = 0.9), 
      # 放大图例字体 
      legend.text  = element_text(size=30),     # 图例里文字的大小
      legend.title = element_text(size=30),     # 图例标题的大小
      legend.position = "right",
      legend.box      = "vertical",
      # 关键：拉大 facet 之间的间距 
      panel.spacing   = unit(0.5, "cm"),        # 图例间距,增加两个 facet 之间的间距（默认约0.1cm）
      # panel.spacing.x = unit(2, "cm"),        # 水平方向间距（如果是 ncol=2）
      # panel.spacing.y = unit(2, "cm"),        # 垂直方向间距（如果有更多行）
      legend.margin = ggplot2::margin(t = 30),  # 给标题留出位置
      plot.margin = unit(c(0,0,0,0), "pt"),     # 零边距
      text = element_text(family = "sans")      # 全局指定支持符号的字体！
    )
  # 3.零间距拼接
  p_final <- p1 +
    inset_element(
      cluster,
      left     = 0.495,
      bottom   = 0.000,
      right    = 0.505,
      top      = 1.000,
      align_to = "panel"
    ) +
    # 用plot_annotation把标题放在最外层，和图例层不冲突
    plot_annotation(
      title = paste0("Cancer: ", cancer1),
      theme = theme(
        plot.title = element_text(
          hjust = 0.995,  # 右对齐，刚好和图例对齐
          vjust = -27.0,  # 上下微调，贴在图例顶部
          size  = 35,
          face  = "bold"
        )
      )
    )
  print(p_final)
  
  # 【核心：自动计算图片宽度，完美解决压缩】
  # 1. 计算最长通路名的字符数
  max_char <- max(nchar(as.character(plotdata$Description)))
  # 2. 自动计算总宽度（主图固定 12 英寸 + Y轴文字动态宽度）
  auto_width <- 7 + (max_char / 12)   
  
  # 保存：使用动态宽度
  ggsave(paste0("final_enrichment_plot-", cancer1, ".png"), 
         plot = p_final, 
         width  = auto_width,   # ✅ 自动变宽，不压缩
         height = 9, 
         dpi    = 300,
         bg     = "white")
  
}
#### （3.1）富集结果绘制 ： 正文表格 --------------------------------- ----
summary_table <- all_enrich %>%
  filter(Padj < 0.05) %>%
  group_by(cancer, by, kappa_type, type) %>%
  summarise(
    显著通路数 = n(),
    fold范围   = paste(round(min(fold),1), "~", round(max(fold),1)),
    代表性通路 = paste(head(Description, 3), collapse = "; "),
    .groups = 'drop'
  ) %>%
  arrange(cancer, by, kappa_type, type)
print(summary_table)
library(openxlsx)
write.xlsx(summary_table, "附表1_富集分析完整统计.xlsx", 
           rowNames = FALSE,
           colNames = TRUE)


# 按显著性水平划分的完整堆叠图（英文版）-附录
p_supp_figure <- all_enrich %>%
  filter(Padj < 0.05) %>%
  mutate(
    significance_level = case_when(
      Padj < 1e-10 ~ "Highly significant",
      Padj < 1e-5 ~ "Moderately significant",
      Padj < 1e-3 ~ "Significant",
      TRUE ~ "Nominally significant"
    ),
    significance_level = factor(significance_level,
                                levels = c("Highly significant",
                                           "Moderately significant",
                                           "Significant",
                                           "Nominally significant"))
  ) %>%
  group_by(cancer, type, aim, by, kappa_type, significance_level) %>%
  summarise(pathway_count = n(), .groups = 'drop') %>%
  mutate(group = paste(type, aim, kappa_type, sep = "_")) %>%
  ggplot(aes(x = pathway_count, y = group, fill = significance_level)) +
  geom_col() +
  facet_grid(cancer ~ by, scales = "free_y", space = "free_y") +
  scale_fill_manual(
    name = "Significance level",
    values = c("Highly significant" = "#CC3333",
               "Moderately significant" = "#E69F00",
               "Significant" = "#6699CC",
               "Nominally significant" = "#999999")
  ) +
  theme_bw() +
  theme(
    axis.text.y = element_text(size = 6),
    strip.text = element_text(size = 10, face = "bold"),
    legend.position = "bottom",
    legend.title = element_text(size = 10),
    legend.text = element_text(size = 8)
  ) +
  labs(
    x = "Number of significant pathways (Padj < 0.05)",
    y = NULL,
    title = ""
  )

#### （3.2）富集结果绘制 ： 正文 通路热图：泛癌共性通路 -------------------------- ----
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(patchwork)

# 读取数据
df <- read_excel("富集通路-new1.xlsx", sheet = 1)

logp_max <- df %>%
  filter(Padj < 0.05, Count >= 3) %>%  # 保持和你最终筛选一致
  group_by(type, cancer, Level_2_Category, Level_1_Category) %>%
  summarise(min_padj = min(Padj, na.rm = TRUE), .groups = "drop") %>%
  mutate(logp = -log10(ifelse(is.na(min_padj), 1, min_padj))) %>%
  pull(logp) %>%
  max(na.rm = TRUE)
# 数据预处理（修复标签逻辑）
heatmap_data_all <- df %>%
  filter(Padj < 0.05, Count >= 3, by == "kegg") %>%
  # 按分组统计【总通路数】（合并所有kappa_type）
  group_by(type, cancer, Level_2_Category, Level_1_Category) %>%
  summarise(
    n_pathways = n_distinct(Description),  # 总数量
    min_padj   = min(Padj, na.rm = TRUE),
    .groups    = "drop") %>%
  # 构建标签：只显示数字
  mutate(
    label_text       = n_pathways,
    logp_Theoretical = -log10(ifelse(is.na(min_padj), 1, min_padj))) %>%
  group_by(Level_2_Category) %>%
  filter(sum(n_pathways) > 0) %>%
  ungroup()
heatmap_gsgc  <- heatmap_data_all %>% filter(type == "gsgcs")
heatmap_coxcs <- heatmap_data_all %>% filter(type == "coxcs")
# 保证kegg和hallmark的图一致
# logp_max      <- max(heatmap_data_all$logp_Theoretical, na.rm = TRUE)


# 排序
category_order <- heatmap_data_all %>%
  arrange(Level_1_Category, Level_2_Category) %>%
  pull(Level_2_Category) %>%
  unique()
# 1. Level 1 Category 列（紧贴热图）
p_level1 <- heatmap_gsgc %>%
  distinct(Level_2_Category, Level_1_Category) %>%
  mutate(
    Level_2_Category = factor(Level_2_Category, levels = category_order),
    x_var = factor("Level 1", levels = "Level 1")) %>%
  ggplot(aes(x = x_var, y = Level_2_Category)) +
  geom_tile(aes(fill = Level_1_Category), color = "white", linewidth = 0.5) +
  scale_fill_manual(
    values = scales::brewer_pal(palette = "Set3")(length(unique(heatmap_data_all$Level_1_Category))),
    name = "Level 1 Category (kegg)") +
  theme_minimal() +
  labs(x = NULL, y = NULL) +
  theme(
    axis.text.x     = element_blank(),
    axis.text.y     = element_text(size = 10),
    axis.title      = element_blank(),
    axis.title.x    = element_text(size = 10),#, face = "bold"),# 
    legend.text     = element_text(size = 10),                 # 图例标签
    legend.title    = element_text(size = 12),#, face = "bold"),  # 图例标题
    panel.grid      = element_blank(),
    legend.position = "none",
    plot.margin     = ggplot2::margin(0, 0, 0, 0, "pt")
  )

# 2. GSGC 热图
p_gsgc <- heatmap_gsgc %>%
  mutate(Level_2_Category = factor(Level_2_Category, levels = category_order)) %>%
  ggplot(aes(x = cancer, y = Level_2_Category)) +
  geom_tile(aes(fill = logp_Theoretical), color = "white", linewidth = 0.5) +
  geom_text(aes(label = label_text), size = 3, color = "black") +
  scale_fill_gradientn(
    colours = c("white", "#1A5592", "#B83D3D"),
    limits = c(0, logp_max),
    name = "-log10(p.adjust)") +
  theme_minimal() +
  labs(title = "GSGCS", x = "Cancer", y = NULL) +
  theme(
    axis.text.x     = element_text(angle = 15, hjust = 1, size = 10),
    axis.text.y     = element_blank(),
    axis.title.y    = element_blank(),
    axis.title.x    = element_text(size = 10),#, face = "bold"),
    legend.text     = element_text(size = 10),                 # 图例标签
    legend.title    = element_text(size = 12),#, face = "bold"),  # 图例标题
    plot.title      = element_text(hjust = 0.5,  size = 12), # face = "bold",
    legend.position = "none",
    panel.grid      = element_blank(),
    plot.margin     = ggplot2::margin(0, 0, 0, 0)
  )
# 3. CoxCS 热图
p_coxcs <- heatmap_coxcs %>%
  mutate(Level_2_Category = factor(Level_2_Category, levels = category_order)) %>%
  ggplot(aes(x = cancer, y = Level_2_Category)) +
  geom_tile(aes(fill = logp_Theoretical), color = "white", linewidth = 0.5) +
  geom_text(aes(label = label_text), size = 3, color = "black") +
  scale_fill_gradientn(
    colours = c("white", "#1A5592", "#B83D3D"),
    limits = c(0, logp_max),
    name = "-log10(p.adjust)") +
  theme_minimal() +
  labs(title = "CoxCS", x = "Cancer", y = NULL) +
  theme(
    axis.text.x     = element_text(angle = 15, hjust = 1, size = 10),
    axis.text.y     = element_blank(),
    axis.title.y    = element_blank(),
    axis.title.x    = element_text(size = 10),#, face = "bold"),
    legend.text     = element_text(size = 10),                 # 图例标签
    legend.title    = element_text(size = 12),#, face = "bold"),  # 图例标题
    plot.title      = element_text(hjust = 0.5, size = 12),#  face = "bold",
    legend.position = "right",
    panel.grid      = element_blank(),
    plot.margin     = ggplot2::margin(0, 0, 0, 0)
  )
# 组合：Level1列宽度再缩窄一半（0.3 → 0.15）
final_plot <- p_level1 + p_gsgc + p_coxcs +
  plot_layout(widths = c(0.15, 3.5, 3.5),guides = "collect") +
  plot_annotation(
    title = "                      ",
    theme = theme(plot.title = element_text(hjust = 0.5, size = 10))) & # , face = "bold"
  theme(legend.position = "right") # theme(legend.position = "bottom")
# 保存
ggsave("新图2/Figure_C_Final_kegg-new.png",final_plot,width = 14, height = 10, dpi = 300)
print(final_plot)


# 数据预处理（修复标签逻辑）
logp_max <- df %>%
  filter(Padj < 0.05, Count >= 3) %>%  # 保持和你最终筛选一致
  group_by(type, cancer, Level_2_Category, Level_1_Category) %>%
  summarise(min_padj = min(Padj, na.rm = TRUE), .groups = "drop") %>%
  mutate(logp = -log10(ifelse(is.na(min_padj), 1, min_padj))) %>%
  pull(logp) %>%
  max(na.rm = TRUE)

heatmap_data_all <- df %>%
  filter(Padj < 0.05, Count >= 3, by == "hallmark") %>%
  # 按分组统计【总通路数】（合并所有kappa_type）
  group_by(type, cancer, Level_2_Category, Level_1_Category) %>%
  summarise(
    n_pathways = n_distinct(Description),  # 总数量
    min_padj   = min(Padj, na.rm = TRUE),
    .groups    = "drop") %>%
  # 构建标签：只显示数字
  mutate(label_text = n_pathways,
         logp_Theoretical = -log10(ifelse(is.na(min_padj), 1, min_padj))) %>%
  group_by(Level_2_Category) %>%
  filter(sum(n_pathways) > 0) %>%
  ungroup()
heatmap_gsgc  <- heatmap_data_all %>% filter(type == "gsgcs")
heatmap_coxcs <- heatmap_data_all %>% filter(type == "coxcs")
# 保证kegg和hallmark的图一致
# logp_max      <- max(heatmap_data_all$logp_Theoretical, na.rm = TRUE)

# 排序
category_order <- heatmap_data_all %>%
  arrange(Level_1_Category, Level_2_Category) %>%
  pull(Level_2_Category) %>%
  unique()
# 1. Level 1 Category 列（紧贴热图）
p_level1 <- heatmap_gsgc %>%
  distinct(Level_2_Category, Level_1_Category) %>%
  mutate(
    Level_2_Category = factor(Level_2_Category, levels = category_order),
    x_var = factor("Level 1", levels = "Level 1")) %>%
  ggplot(aes(x = x_var, y = Level_2_Category)) +
  geom_tile(aes(fill = Level_1_Category), color = "white", linewidth = 0.5) +
  scale_fill_manual(
    values = scales::brewer_pal(palette = "Set3")(length(unique(heatmap_data_all$Level_1_Category))),
    name = "Level 1 Category (hallmark)") +
  theme_minimal() +
  labs(x = NULL, y = NULL) +
  theme(
    axis.text.x     = element_blank(),
    axis.text.y     = element_text(size = 10),
    axis.title      = element_blank(),
    axis.title.x    = element_text(size = 10),#, face = "bold"),
    legend.text     = element_text(size = 10),                 # 图例标签
    legend.title    = element_text(size = 12),#, face = "bold"),  # 图例标题
    panel.grid      = element_blank(),
    legend.position = "none",
    plot.margin     = ggplot2::margin(0, 0, 0, 0)
  )

# 2. GSGC 热图
p_gsgc <- heatmap_gsgc %>%
  mutate(Level_2_Category = factor(Level_2_Category, levels = category_order)) %>%
  ggplot(aes(x = cancer, y = Level_2_Category)) +
  geom_tile(aes(fill = logp_Theoretical), color = "white", linewidth = 0.5) +
  geom_text(aes(label = label_text), size = 3, color = "black") +
  scale_fill_gradientn(
    colours = c("white", "#1A5592", "#B83D3D"),
    limits = c(0, logp_max),
    name = "-log10(p.adjust)") +
  theme_minimal() +
  labs(title = "GSGCS", x = "Cancer", y = NULL) +
  theme(
    axis.text.x     = element_text(angle = 15, hjust = 1, size = 10),
    axis.text.y     = element_blank(),
    axis.title.y    = element_blank(),
    axis.title.x    = element_text(size = 10),#, face = "bold"),
    legend.text     = element_text(size = 10),                 # 图例标签
    legend.title    = element_text(size = 12),#, face = "bold"),  # 图例标题
    plot.title      = element_text(hjust = 0.5, size = 12), # face = "bold",
    legend.position = "none",
    panel.grid      = element_blank(),
    plot.margin     = ggplot2::margin(0, 0, 0, 0)
  )
# 3. CoxCS 热图
p_coxcs <- heatmap_coxcs %>%
  mutate(Level_2_Category = factor(Level_2_Category, levels = category_order)) %>%
  ggplot(aes(x = cancer, y = Level_2_Category)) +
  geom_tile(aes(fill = logp_Theoretical), color = "white", linewidth = 0.5) +
  geom_text(aes(label = label_text), size = 3, color = "black") +
  scale_fill_gradientn(
    colours = c("white", "#1A5592", "#B83D3D"),
    limits = c(0, logp_max),
    name = "-log10(p.adjust)") +
  theme_minimal() +
  labs(title = "CoxCS", x = "Cancer", y = NULL) +
  theme(
    axis.text.x     = element_text(angle = 15, hjust = 1, size = 10),
    axis.text.y     = element_blank(),
    axis.title.y    = element_blank(),
    axis.title.x    = element_text(size = 10),#, face = "bold"),
    legend.text     = element_text(size = 10),                 # 图例标签
    legend.title    = element_text(size = 12),#, face = "bold"),  # 图例标题
    plot.title      = element_text(hjust = 0.5, size = 12), # face = "bold",
    legend.position = "right",
    panel.grid      = element_blank(),
    plot.margin     = ggplot2::margin(0, 0, 0, 0)
  )
# 组合：Level1列宽度再缩窄一半（0.3 → 0.15）
final_plot <- p_level1 + p_gsgc + p_coxcs +
  plot_layout(widths = c(0.15, 3.5, 3.5),guides = "collect") +
  plot_annotation(title = "                     ",
                  theme = theme(plot.title = element_text(hjust = 0.5, size = 10))) & #, face = "bold"))) &
  theme(legend.position = "right") # theme(legend.position = "bottom")
# 保存
ggsave("新图2/Figure_C_Final_hallmark-new.png",final_plot,width = 14, height = 10, dpi = 300)
print(final_plot)



#### （3.3）富集结果绘制 ： 正文 通路富集图 -------------------------- ----
library(ggplot2)
library(dplyr)
library(patchwork)
library(ggnewscale)
library(RColorBrewer)

# 准备数据：取每个组合下 fold 前 5 的通路
pdata_top5 <- all_enrich %>%
  filter(Padj < 0.05) %>%
  group_by(cancer, type, aim, by, kappa_type) %>%
  arrange(desc(fold), .by_group = TRUE) %>%  # 按 fold 降序排列
  slice_head(n = 10) %>%                      # 取前 5
  ungroup() %>%
  # 创建用于排序的通路名称（保证每个癌症内通路按 fold 排序）
  group_by(cancer, by) %>%
  mutate(
    Description = factor(Description, levels = unique(Description[order(-fold)]))
  ) %>%
  ungroup()

# 拆分 hallmark 和 kegg 数据
pdata_hallmark <- pdata_top5 %>% filter(by == "hallmark"&  kappa_type == "Theoretical")
pdata_kegg     <- pdata_top5 %>% filter(by == "kegg"&      kappa_type == "Theoretical")

# 绘制 Hallmark 图
if(nrow(pdata_hallmark) > 0) {
  # 设置因子水平，确保横轴通路按 fold 排序
  pdata_hallmark <- pdata_hallmark %>%
    group_by(cancer) %>%
    mutate(
      Description = factor(Description, levels = unique(Description[order(fold, decreasing = TRUE)]))
    ) %>%
    ungroup()
  
  p_hallmark <- ggplot(pdata_hallmark, aes(x = Description, y = cancer)) +
    geom_point(aes(color = fold, size = LOGP, shape = type), alpha = 0.85, stroke = 0.8) +
    facet_grid(aim ~ ., scales = "free_y", space = "free_y") +  # 行为 aim，列为空
    # scale_color_gradientn(
    #   low = "#6699CC",      # 低 fold 颜色（蓝色）
    #   high = "#CC3333",     # 高 fold 颜色（红色）
    #   name = "Fold enrichment"
    # ) +
    scale_color_gradientn(
      colours = brewer.pal(11, "Spectral"),
      name = "Fold enrichment"
    ) +
    scale_size_continuous(name = "-log10(Padj)", range = c(2, 10)) +
    # scale_shape_manual(values = c("gsgc" = 16, "coxcs" = 17), name = "Method") +
    scale_shape_manual(
      values = c("gsgc" = 16, "coxcs" = 17), 
      name = "Method",
      labels = c("gsgc" = "GSGCS", "coxcs" = "CoxCS")  # 只改显示标签，不改数据
    )+
    scale_x_discrete(position = "bottom") +
    labs(
      title = "Hallmark pathways - Top 15",
      x = NULL,
      y = "Cancer type"
    ) +
    theme_bw() +
    theme(
      axis.text.x      = element_text(angle = 75, hjust = 1, size = 10),
      axis.text.y      = element_text(size = 10),
      strip.text       = element_text(size = 10, face = "bold"),
      legend.position  = "right",
      legend.box       = "vertical",
      legend.text      = element_text(size = 10),
      legend.title     = element_text(size = 10),
      plot.title       = element_text(hjust = 0.5, size = 10, face = "bold"),
      panel.grid.major = element_line(color = "gray90"),
      panel.grid.minor = element_blank()
    )
  
  print(p_hallmark)
  ggsave("Figure_Hallmark_Top10_Fold_Horizontal-new.png", p_hallmark, width = 16, height = 8, dpi = 300)
}
# 绘制 KEGG 图 
if(nrow(pdata_kegg) > 0) {
  pdata_kegg <- pdata_kegg %>%
    group_by(cancer) %>%
    mutate(
      Description = factor(Description, levels = unique(Description[order(fold, decreasing = TRUE)]))
    ) %>%
    ungroup()
  
  p_kegg <- ggplot(pdata_kegg, aes(x = Description, y = cancer)) +
    geom_point(aes(color = fold, size = LOGP, shape = type), alpha = 0.85, stroke = 0.8) +
    facet_grid(aim ~ ., scales = "free_y", space = "free_y") +
    scale_color_gradientn(
      colours = brewer.pal(11, "Spectral"),
      name = "Fold enrichment"
    ) +
    scale_size_continuous(name = "-log10(Padj)", range = c(2, 10)) +
    # scale_shape_manual(values = c("gsgc" = 16, "coxcs" = 17), name = "Method") +
    scale_shape_manual(
      values = c("gsgc" = 16, "coxcs" = 17), 
      name = "Method",
      labels = c("gsgc" = "GSGCS", "coxcs" = "CoxCS")  # 只改显示标签，不改数据
    )+
    scale_x_discrete(position = "bottom") +
    labs(
      title = "KEGG pathways - Top 15",
      x = NULL,
      y = "Cancer type"
    ) +
    theme_bw() +
    theme(
      axis.text.x = element_text(angle = 75, hjust = 1, size = 10),
      axis.text.y = element_text(size = 10),
      strip.text = element_text(size = 10, face = "bold"),
      legend.position = "right",
      legend.box = "vertical",
      legend.text = element_text(size = 10),
      legend.title = element_text(size = 10),
      plot.title = element_text(hjust = 0.5, size = 10, face = "bold"),
      panel.grid.major = element_line(color = "gray90"),
      panel.grid.minor = element_blank()
    )
  
  print(p_kegg)
  ggsave("Figure_KEGG_Top10_Fold_Horizontal-new.png", p_kegg, width = 16, height = 8, dpi = 300)
}


# 整理通路结果
# > length(all_enrich$Description)
# [1] 5183
# > length(unique(all_enrich$Description))
# [1] 375
pathway = all_enrich # %>%dplyr::select(cancer,type,aim,kappa_type,by,Description,ID,Count,geneID)
write.xlsx(pathway, "富集通路-new1.xlsx",
           rowNames = FALSE,
           colNames = TRUE)


#### （3.4）富集结果绘制 ： Figure 3 绘制哑铃图 ----
library(ggplot2)
library(dplyr)
library(tidyr)
library(patchwork)

# 1. Complete data  
dat <- tribble(
  ~Scheme,       ~Source, ~Pathway,  ~Cancer, ~Outcome, ~Method, ~Reported, ~Detected, ~Rate,
  # ---------------- Scheme one: TCGA, KEGG 
  "Scheme one",  "TCGA",  "KEGG",     "BRCA", "OS",      "CoxCS", 13,  7, 53.85,
  "Scheme one",  "TCGA",  "KEGG",     "BRCA", "OS",      "GSGCS", 13, 11, 84.62,
  "Scheme one",  "TCGA",  "KEGG",     "BRCA", "PFI",     "CoxCS", 13,  6, 46.15,
  "Scheme one",  "TCGA",  "KEGG",     "BRCA", "PFI",     "GSGCS", 13, 11, 84.62,
  "Scheme one",  "TCGA",  "KEGG",     "LUAD", "OS",      "CoxCS", 68, 40, 58.82,
  "Scheme one",  "TCGA",  "KEGG",     "LUAD", "OS",      "GSGCS", 68, 42, 61.76,
  "Scheme one",  "TCGA",  "KEGG",     "LUAD", "PFI",     "CoxCS", 68, 22, 32.35,
  "Scheme one",  "TCGA",  "KEGG",     "LUAD", "PFI",     "GSGCS", 68, 31, 45.59,
  "Scheme one",  "TCGA",  "KEGG",     "OV",   "OS",      "CoxCS", 12,  1,  8.33,
  "Scheme one",  "TCGA",  "KEGG",     "OV",   "OS",      "GSGCS", 12,  6, 50.00,
  "Scheme one",  "TCGA",  "KEGG",     "OV",   "PFI",     "CoxCS", 12,  0,  0.00,
  "Scheme one",  "TCGA",  "KEGG",     "OV",   "PFI",     "GSGCS", 12,  6, 50.00,
  "Scheme one",  "TCGA",  "KEGG",     "STAD", "OS",      "CoxCS", 80, 17, 21.25,
  "Scheme one",  "TCGA",  "KEGG",     "STAD", "OS",      "GSGCS", 80, 56, 70.00,
  "Scheme one",  "TCGA",  "KEGG",     "STAD", "PFI",     "CoxCS", 80, 19, 23.75,
  "Scheme one",  "TCGA",  "KEGG",     "STAD", "PFI",     "GSGCS", 80, 65, 81.25,
  "Scheme one",  "TCGA",  "KEGG",     "THYM", "OS",      "CoxCS", 11,  7, 63.64,
  "Scheme one",  "TCGA",  "KEGG",     "THYM", "OS",      "GSGCS", 11,  4, 36.36,
  "Scheme one",  "TCGA",  "KEGG",     "THYM", "PFI",     "CoxCS", 11,  2, 18.18,
  "Scheme one",  "TCGA",  "KEGG",     "THYM", "PFI",     "GSGCS", 11,  9, 81.82,
  # ------------- Scheme one: TCGA, Hallmark 
  "Scheme one",  "TCGA",  "Hallmark", "BRCA", "OS",      "CoxCS", 11,  6, 54.55,
  "Scheme one",  "TCGA",  "Hallmark", "BRCA", "OS",      "GSGCS", 11,  6, 54.55,
  "Scheme one",  "TCGA",  "Hallmark", "BRCA", "PFI",     "CoxCS", 11,  4, 36.36,
  "Scheme one",  "TCGA",  "Hallmark", "BRCA", "PFI",     "GSGCS", 11,  6, 54.55,
  "Scheme one",  "TCGA",  "Hallmark", "OV",   "OS",      "CoxCS",  8,  0,  0.00,
  "Scheme one",  "TCGA",  "Hallmark", "OV",   "OS",      "GSGCS",  8,  4, 50.00,
  "Scheme one",  "TCGA",  "Hallmark", "OV",   "PFI",     "CoxCS",  8,  0,  0.00,
  "Scheme one",  "TCGA",  "Hallmark", "OV",   "PFI",     "GSGCS",  8,  4, 50.00,
  "Scheme one",  "TCGA",  "Hallmark", "STAD", "OS",      "CoxCS", 75, 17, 22.67,
  "Scheme one",  "TCGA",  "Hallmark", "STAD", "OS",      "GSGCS", 75, 66, 88.00,
  "Scheme one",  "TCGA",  "Hallmark", "STAD", "PFI",     "CoxCS", 75, 21, 28.00,
  "Scheme one",  "TCGA",  "Hallmark", "STAD", "PFI",     "GSGCS", 75, 66, 88.00,
  # ---------------- Scheme one: PCAWG 
  "Scheme one",  "PCAWG", "KEGG",     "BRCA", "Time",    "CoxCS", 14,  5, 35.71,
  "Scheme one",  "PCAWG", "KEGG",     "BRCA", "Time",    "GSGCS", 14,  5, 35.71,
  "Scheme one",  "PCAWG", "KEGG",     "LUAD", "Time",    "CoxCS", 68, 11, 16.18,
  "Scheme one",  "PCAWG", "KEGG",     "LUAD", "Time",    "GSGCS", 68,  5,  7.35,
  "Scheme one",  "PCAWG", "KEGG",     "OV",   "Time",    "CoxCS", 12,  0,  0.00,
  "Scheme one",  "PCAWG", "KEGG",     "OV",   "Time",    "GSGCS", 12,  6, 50.00,
  "Scheme one",  "PCAWG", "KEGG",     "STAD", "Time",    "CoxCS", 80, 49, 61.25,
  "Scheme one",  "PCAWG", "KEGG",     "STAD", "Time",    "GSGCS", 80,  0,  0.00,
  "Scheme one",  "PCAWG", "Hallmark", "OV",   "Time",    "CoxCS",  8,  0,  0.00,
  "Scheme one",  "PCAWG", "Hallmark", "OV",   "Time",    "GSGCS",  8,  2, 25.00,
  # ---------------- Scheme two: TCGA Hallmark 
  "Scheme two",  "TCGA",  "Hallmark", "BRCA", "PFI",     "CoxCS", 11,  0,  0.00,
  "Scheme two",  "TCGA",  "Hallmark", "BRCA", "PFI",     "GSGCS", 11,  1,  9.09,
  "Scheme two",  "TCGA",  "Hallmark", "STAD", "PFI",     "CoxCS", 75,  0,  0.00,
  "Scheme two",  "TCGA",  "Hallmark", "STAD", "PFI",     "GSGCS", 75,  1,  1.33
)
# 2. Convert to paired wide format 
wide <- dat %>%select(Scheme, Source, Pathway, Cancer,
                      Outcome, Method, Rate) %>%
  pivot_wider(names_from  = Method,
              values_from = Rate) %>%
  mutate(Delta = GSGCS - CoxCS)

# Main figure only uses Scheme one
wide_main <- wide %>% filter(Scheme == "Scheme one")

# 3. Ordering function 
cancer_order <- c(BRCA = 1, LUAD = 2, OV= 3, STAD = 4, THYM = 5)
wide_main <- wide_main %>%
  mutate(pathway_order = case_when(Pathway == "KEGG"     ~ 1,
                                   Pathway == "Hallmark" ~ 2,
                                   TRUE ~ 9),
         cancer_order = unname(cancer_order[Cancer]),
         outcome_order = case_when(Outcome == "OS"   ~ 1,
                                   Outcome == "PFI"  ~ 2,
                                   Outcome == "Time" ~ 1,
                                   TRUE ~ 9),
         Analysis = case_when(Source == "TCGA" ~ paste0(Cancer, "\u2013", Outcome," | ", Pathway),
                              Source == "PCAWG" ~ paste0(Cancer," | ", Pathway))
  )
# 4. Plot function 
col_coxcs     <- "#1F77B4"
col_gsgcs     <- "#FF7F0E"
col_connector <- "#B8B8B8"
col_grid      <- "#E6E6E6"
col_text      <- "#222222"

draw_dumbbell_main <- function(dat_source,panel_title) {
  d <- wide_main %>%
    filter(Source == dat_source) %>%
    arrange(pathway_order,cancer_order,outcome_order) %>%
    mutate(Analysis = factor(Analysis,levels = rev(unique(Analysis))))
  ggplot(d,aes(y = Analysis)) +
    # connector
    geom_segment(aes(x = CoxCS,xend = GSGCS,yend = Analysis),linewidth = 0.75,color = col_connector) +
    # CoxCS
    geom_point(aes(x = CoxCS,color = "CoxCS"),size = 3.0) +
    # GSGCS
    geom_point(aes(x = GSGCS,color = "GSGCS"),size = 3.0) +
    # delta
    geom_text(aes(x = 103,label = sprintf("%+.2f", Delta)),hjust = 0,size = 3.0,color = col_text) +
    annotate("text",x = 103,y = Inf,label = "\u0394 (pp)",hjust = 0,vjust = 1.2,size = 3.1,fontface = "bold",color = col_text) +
    scale_color_manual(values = c("CoxCS" = col_coxcs,"GSGCS" = col_gsgcs)) +
    scale_x_continuous(limits = c(0, 116),breaks = seq(0, 100, 20),expand = expansion(mult = c(0.002, 0.002))) +
    labs(title = panel_title,
         x = "Detection rate of previously reported prognostic biomarkers (%)",
         y = NULL,color = NULL) +
    theme_classic(base_size = 10.5) +
    theme(
      plot.title   = element_text(face = "bold",size = 12,hjust = 0,color = col_text,margin = margin(b = 6)),
      axis.text.y  = element_text(size = 8.8,color = col_text,margin = margin(r = 6)),
      axis.text.x  = element_text(size = 8.8,color = col_text),
      axis.title.x = element_text(size = 9.5,color = col_text,margin = margin(t = 7)),
      panel.grid.major.x = element_line(color = col_grid,linewidth = 0.35),
      panel.grid.minor = element_blank(),
      axis.line.y      = element_blank(),
      axis.ticks.y     = element_blank(),
      legend.position  = "bottom",
      legend.direction = "horizontal",
      legend.text = element_text(size = 9.2),
      legend.key  = element_blank(),
      plot.margin = margin(t = 5,r = 38,b = 5,l = 5))
}

# 5. Draw two panels 
p_tcga  <- draw_dumbbell_main(dat_source = "TCGA", panel_title = "A  TCGA")
p_pcawg <- draw_dumbbell_main(dat_source = "PCAWG",panel_title = "B  PCAWG")

# 6. Combine 
fig3 <- wrap_plots(p_tcga,p_pcawg,ncol = 1,heights = c(3.1, 1.15),guides = "collect") &theme(legend.position = "bottom")
fig3

# 7. Export 
ggsave(filename = "C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/文章图选择/改为GSGCS/Figure3_BIB_dumbbell.pdf", 
       plot = fig3,width = 7.5,height = 9.0,units = "in",device = cairo_pdf)
ggsave(filename = "C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/文章图选择/改为GSGCS/Figure3_BIB_dumbbell.tiff",
       plot = fig3,width = 7.5,height = 9.0,units = "in",dpi = 600,compression = "lzw")
ggsave(filename = "C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/文章图选择/改为GSGCS/Figure3_BIB_dumbbell.jpg",
       plot = fig3,width = 7.5,height = 9.0,units = "in",dpi = 600)

# 附录图
wide_s2 <- wide %>%
  filter(Scheme == "Scheme two") %>%
  mutate(Analysis = paste0(Cancer,"\u2013",Outcome," | ",Pathway),
         cancer_order = unname(cancer_order[Cancer])) %>%
  arrange(cancer_order,Outcome) %>%
  mutate(Analysis = factor(Analysis,levels = rev(unique(Analysis))))

figS_scheme2 <- ggplot(wide_s2,aes(y = Analysis)) +
  geom_segment(aes(x = CoxCS,xend = GSGCS,yend = Analysis),linewidth = 0.75,color = col_connector) +
  geom_point(  aes(x = CoxCS,color = "CoxCS"),size = 3.0) +
  geom_point(  aes(x = GSGCS,color = "GSGCS"),size = 3.0) +
  geom_text(   aes(x = 10.5,label = sprintf("%+.2f", Delta)),hjust = 0,size = 3.0,color = col_text) +
  annotate("text", x = 10.5,y = Inf,label = "\u0394 (pp)",hjust = 0,vjust = 1.2,size = 3.1,fontface = "bold",color = col_text) +
  scale_color_manual(values = c("CoxCS" = col_coxcs,"GSGCS" = col_gsgcs)) +
  scale_x_continuous(limits = c(0, 13),breaks = c(0, 2, 4, 6, 8, 10),expand = expansion(mult = c(0.002, 0.002))) +
  labs(title = "Secondary screening scheme",
       subtitle = "TCGA Hallmark analyses",
       x = "Detection rate of previously reported prognostic biomarkers (%)",
       y = NULL,color = NULL) +
  theme_classic(base_size = 10.5) +
  theme(
    plot.title    = element_text(face = "bold",size = 11.5,hjust = 0,color = col_text),
    plot.subtitle = element_text(size = 9.5,hjust = 0,color = "#555555",margin = margin(b = 6)),
    axis.text.y   = element_text(size = 9,color = col_text),
    axis.text.x   = element_text(size = 8.8,color = col_text),
    axis.title.x  = element_text(size = 9.5,color = col_text,margin = margin(t = 7)),
    panel.grid.major.x = element_line(color = col_grid,linewidth = 0.35),
    panel.grid.minor   = element_blank(),
    axis.line.y        = element_blank(),
    axis.ticks.y       = element_blank(),
    legend.position    = "bottom",
    legend.direction   = "horizontal",
    legend.text        = element_text(size = 9.2),
    legend.key         = element_blank(),
    plot.margin = margin(t = 6,r = 38,b = 6,l = 6))
figS_scheme2
# Export supplementary figure 
ggsave(filename = "C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/文章图选择/改为GSGCS/FigureS_scheme2_dumbbell-更正.pdf", 
       plot = figS_scheme2,width = 6.5,height = 4.0,units = "in",device = cairo_pdf)
ggsave(filename = "C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/文章图选择/改为GSGCS/FigureS_scheme2_dumbbell-更正.tiff",
       plot = figS_scheme2,width = 6.5,height = 4.0,units = "in",dpi = 600,compression = "lzw")
ggsave(filename = "C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/文章图选择/改为GSGCS/FigureS_scheme2_dumbbell-更正.jpg",
       plot = figS_scheme2,width = 6.5,height = 4.0,units = "in")


#### （3.5）富集结果绘制 ： figure4 A-B -----
library(dplyr)
library(tidyr)
library(ggplot2)
library(patchwork)
library(ComplexUpset)
library(readxl)
library(igraph)
library(ggraph)
setwd("C:/Users/hp/Desktop/课题组/其他事务/250827-贺老师-服务器&数据/全场景实验结果/truedata/260428结果/")


library(ggplot2)
library(dplyr)
library(tidyr)
library(patchwork)
library(RColorBrewer)
library(scales)

# 读取数据
df <- read_excel("富集通路-new1.xlsx", sheet = 1)
all_enrich= df

# 准备数据：取每个组合下 fold 前 5 的通路
pdata_top5 <- all_enrich %>%
  filter(Padj < 0.05) %>%
  group_by(cancer, type, aim, by, kappa_type) %>%
  arrange(desc(fold), .by_group = TRUE) %>%  # 按 fold 降序排列
  slice_head(n = 10) %>%                      # 取前 5
  ungroup() %>%
  # 创建用于排序的通路名称（保证每个癌症内通路按 fold 排序）
  group_by(cancer, by) %>%
  mutate(
    Description = factor(Description, levels = unique(Description[order(-fold)]))
  ) %>%
  ungroup()

# 拆分 hallmark 和 kegg 数据
pdata_hallmark <- pdata_top5 %>% filter(by == "hallmark"&  kappa_type == "Theoretical")
pdata_kegg     <- pdata_top5 %>% filter(by == "kegg"&      kappa_type == "Theoretical")

# Figure 4A: Hallmark
# 上：GSGCS vs CoxCS 数量 summary strip
# 下：左右分栏 paired bubble plot
#
# 重要：
# 只使用已经生成好的 pdata_hallmark
# 不改变前面的富集分析、Top10筛选、fold排序等计算逻辑

if(nrow(pdata_hallmark) > 0) {
  
  # 0. 基础设置
  # 最终图从上到下：
  # BRCA -> LUAD -> OV -> STAD -> THYM
  #
  # ggplot 离散 y 轴从下向上排列，因此 factor levels 反向设置
  cancer_display <- c("BRCA", "LUAD", "OV", "STAD", "THYM")
  cancer_levels  <- rev(cancer_display)
  
  aim_levels    <- c("os", "pfi", "time")
  method_levels <- c("gsgcs", "coxcs")
  
  aim_labels <- c(
    os   = "OS",
    pfi  = "PFI",
    time = "Time"
  )
  
  method_labels <- c(
    gsgcs = "GSGCS",
    coxcs = "CoxCS"
  )
  
  # PART A
  # Number of significant enriched pathways
  #
  # 重要：
  # 这里必须使用 all_enrich，
  # 不能使用 pdata_hallmark，
  # 因为 pdata_hallmark 已经经过 Top-ranked 截断
  
  # 1. 从全部富集结果中提取 Hallmark
  
  
  hall_all <- all_enrich %>%
    mutate(
      cancer     = toupper(trimws(as.character(cancer))),
      aim        = tolower(trimws(as.character(aim))),
      type       = tolower(trimws(as.character(type))),
      by         = tolower(trimws(as.character(by))),
      kappa_type = trimws(as.character(kappa_type))
    ) %>%
    mutate(
      type = dplyr::recode(
        type,
        "gsgc"  = "gsgcs",
        "gsgcs" = "gsgcs"
      )
    ) %>%
    filter(
      by == "hallmark",
      kappa_type == "Theoretical",
      Padj < 0.05
    )
  
  
  # 检查方法名称
  bad_methods <- setdiff(
    unique(na.omit(hall_all$type)),
    method_levels
  )
  
  if(length(bad_methods) > 0) {
    stop(
      paste0(
        "Hallmark 数据中存在无法识别的方法名称：",
        paste(bad_methods, collapse = ", "),
        "\n允许值只能是：gsgcs, coxcs"
      )
    )
  }
  
  
  hall_all <- hall_all %>%
    mutate(
      cancer = factor(cancer, levels = cancer_levels),
      aim    = factor(aim, levels = aim_levels),
      type   = factor(type, levels = method_levels)
    )
  
  # 2. 统计全部显著 pathway 数量
  hall_count <- hall_all %>%
    filter(
      !is.na(cancer),
      !is.na(aim),
      !is.na(type),
      !is.na(Description)
    ) %>%
    distinct(cancer, aim, type, Description) %>%
    count(cancer, aim, type, name = "N") %>%
    complete(
      cancer = factor(cancer_levels, levels = cancer_levels),
      aim    = factor(aim_levels, levels = aim_levels),
      type   = factor(method_levels, levels = method_levels),
      fill   = list(N = 0)
    ) %>%
    mutate(
      N_plot = ifelse(type == "gsgcs", -N, N)
    )
  
  
  # 3. Delta = GSGCS - CoxCS
  hall_delta <- hall_count %>%
    dplyr::select(cancer, aim, type, N) %>%
    pivot_wider(
      names_from  = type,
      values_from = N,
      values_fill = 0
    ) %>%
    mutate(
      delta = gsgcs - coxcs,
      
      delta_label = case_when(
        delta > 0 ~ paste0("+", delta),
        TRUE      ~ as.character(delta)
      ),
      
      delta_group = case_when(
        delta > 0 ~ "positive",
        delta < 0 ~ "negative",
        TRUE      ~ "zero"
      )
    )
  
  # 4. 上图左右严格对称
  hall_maxN <- max(hall_count$N, na.rm = TRUE)
  
  if(!is.finite(hall_maxN) || hall_maxN <= 0) {
    hall_maxN <- 1
  }
  
  # 为柱末端数字留空间
  hall_lim <- hall_maxN * 1.20
  
  # 5. mirrored bar plot
  p_hall_bar <- ggplot(
    hall_count,
    aes(x = N_plot, y = cancer, fill = type)
  ) +
    
    geom_col(
      width = 0.52,
      na.rm = TRUE
    ) +
    
    geom_vline(
      xintercept = 0,
      color = "grey25",
      linewidth = 0.45
    ) +
    
    geom_text(
      aes(
        label = N,
        hjust = ifelse(type == "gsgcs", 1.15, -0.15)
      ),
      size = 3,
      color = "grey20",
      na.rm = TRUE
    ) +
    
    facet_grid(
      aim ~ .,
      scales = "free_y",
      space  = "free_y",
      switch = "y",
      labeller = labeller(
        aim = aim_labels
      )
    ) +
    
    scale_fill_manual(
      values = c(
        gsgcs = "#EF6258",
        coxcs = "#4C89CF"
      ),
      breaks = c("gsgcs", "coxcs"),
      labels = c("GSGCS", "CoxCS"),
      drop   = FALSE,
      name   = NULL
    ) +
    
    scale_x_continuous(
      limits = c(-hall_lim, hall_lim),
      breaks = pretty(
        c(-hall_maxN, hall_maxN),
        n = 5
      ),
      labels = function(x) abs(x),
      expand = expansion(mult = c(0.01, 0.01))
    ) +
    
    labs(
      title = "Number of significant enriched pathways",
      x = "Number of pathways",
      y = NULL
    ) +
    
    theme_bw(base_size = 10) +
    
    theme(
      plot.title = element_text(
        hjust = 0.5,
        face = "bold",
        size = 11,
        margin = margin(b = 5)
      ),
      
      axis.text.y = element_text(size = 9),
      axis.text.x = element_text(size = 8),
      
      strip.background = element_rect(
        fill = "grey90",
        color = "grey45",
        linewidth = 0.4
      ),
      
      strip.text.y.left = element_text(
        angle = 0,
        face = "bold",
        size = 8.5
      ),
      
      panel.grid.major.y = element_blank(),
      panel.grid.minor   = element_blank(),
      
      panel.grid.major.x = element_line(
        color = "grey92",
        linewidth = 0.3
      ),
      
      panel.spacing.y = unit(0.45, "cm"),
      
      legend.position = "top",
      legend.direction = "horizontal",
      legend.justification = "center",
      
      legend.key.width  = unit(0.55, "cm"),
      legend.key.height = unit(0.30, "cm"),
      
      plot.margin = margin(4, 2, 2, 5)
    )
  
  
  # 6. Delta 独立列
  p_hall_delta <- ggplot(
    hall_delta,
    aes(x = 1, y = cancer)
  ) +
    
    geom_text(
      aes(
        label = delta_label,
        color = delta_group
      ),
      size = 3.2,
      fontface = "bold"
    ) +
    
    facet_grid(
      aim ~ .,
      scales = "free_y",
      space = "free_y"
    ) +
    
    scale_color_manual(
      values = c(
        positive = "#EF4B4B",
        negative = "#377FC1",
        zero     = "grey35"
      ),
      guide = "none"
    ) +
    
    scale_x_continuous(
      limits = c(0.65, 1.35),
      expand = c(0, 0)
    ) +
    
    labs(
      title = "\u0394 (GSGCS \u2212 CoxCS)"
    ) +
    
    theme_void(base_size = 10) +
    
    theme(
      plot.title = element_text(
        hjust = 0.5,
        face = "bold",
        size = 10,
        margin = margin(b = 5)
      ),
      
      panel.spacing.y = unit(0.45, "cm"),
      
      strip.text = element_blank(),
      
      plot.margin = margin(4, 5, 2, 0)
    )
  
  # 7. 上半部分组合
  p_hall_top <-
    (p_hall_bar | p_hall_delta) +
    plot_layout(
      widths = c(0.94, 0.06)
    )
  
  # PART B
  # Top-ranked enriched pathways
  #
  # 仍使用 pdata_hallmark
  # 保留原 Top-ranked enrichment profile
  
  # 8. bubble 数据
  
  hall_bubble <- pdata_hallmark %>%
    mutate(
      cancer = toupper(trimws(as.character(cancer))),
      aim    = tolower(trimws(as.character(aim))),
      type   = tolower(trimws(as.character(type)))
    ) %>%
    
    mutate(
      type = dplyr::recode(
        type,
        "gsgc"  = "gsgcs",
        "gsgcs" = "gsgcs"
      )
    ) %>%
    
    filter(
      !is.na(cancer),
      !is.na(aim),
      !is.na(type),
      !is.na(Description),
      !is.na(fold),
      !is.na(LOGP),
      is.finite(fold),
      is.finite(LOGP)
    ) %>%
    
    mutate(
      cancer = factor(cancer, levels = cancer_levels),
      aim    = factor(aim, levels = aim_levels),
      type   = factor(type, levels = method_levels)
    )
  
  
  if(nrow(hall_bubble) == 0) {
    stop(
      "hall_bubble 没有可用于绘图的数据，请检查 fold 和 LOGP。"
    )
  }
  
  # 9. 两方法使用相同 pathway 顺序
  hall_pathway_order <- hall_bubble %>%
    group_by(Description) %>%
    summarise(
      max_fold = max(fold, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    arrange(desc(max_fold)) %>%
    pull(Description)
  
  
  hall_bubble <- hall_bubble %>%
    mutate(
      Description = factor(
        as.character(Description),
        levels = hall_pathway_order
      ),
      
      type = factor(
        as.character(type),
        levels = c("gsgcs", "coxcs")
      )
    )
  
  # 10. 两列共用 Fold enrichment 范围
  
  hall_fold_range <- range(
    hall_bubble$fold,
    na.rm = TRUE
  )
  
  if(
    !all(is.finite(hall_fold_range)) ||
    diff(hall_fold_range) == 0
  ) {
    
    hall_fold_range <- c(
      min(hall_bubble$fold, na.rm = TRUE) - 0.01,
      max(hall_bubble$fold, na.rm = TRUE) + 0.01
    )
  }
  
  # 11. 两列共用 LOGP 范围
  hall_logp_range <- range(
    hall_bubble$LOGP,
    na.rm = TRUE
  )
  
  if(
    !all(is.finite(hall_logp_range)) ||
    diff(hall_logp_range) == 0
  ) {
    
    hall_logp_range <- c(
      max(
        0,
        min(hall_bubble$LOGP, na.rm = TRUE) - 0.01
      ),
      
      max(hall_bubble$LOGP, na.rm = TRUE) + 0.01
    )
  }
  
  # 12. bubble plot
  p_hall_bottom <- ggplot(
    hall_bubble,
    aes(
      x = Description,
      y = cancer
    )
  ) +
    
    geom_point(
      aes(
        color = fold,
        size  = LOGP,
        shape = type
      ),
      alpha = 0.88,
      stroke = 0.6,
      na.rm = TRUE
    ) +
    
    facet_grid(
      rows = vars(aim),
      cols = vars(type),
      scales = "fixed",
      space  = "fixed",
      labeller = labeller(
        aim = aim_labels,
        type = c(
          gsgcs = "GSGCS",
          coxcs = "CoxCS"
        )
      )
    ) +
    
    # Fold enrichment：
    # 反向 Spectral
    scale_color_gradientn(
      colours = rev(
        RColorBrewer::brewer.pal(
          11,
          "Spectral"
        )
      ),
      
      limits = hall_fold_range,
      
      name = "Fold enrichment",
      
      guide = guide_colorbar(
        title.position = "top",
        title.hjust = 0.5,
        direction = "horizontal",
        barwidth  = unit(3.5, "cm"),
        barheight = unit(0.30, "cm"),
        order = 2
      )
    ) +
    
    scale_size_continuous(
      limits = hall_logp_range,
      
      range = c(2, 10),
      
      name = "-log10(Padj)",
      
      guide = guide_legend(
        title.position = "left",
        direction = "horizontal",
        nrow = 1,
        order = 1,
        
        override.aes = list(
          shape = 16,
          color = "black",
          alpha = 1
        )
      )
    ) +
    
    # Method：
    # ● GSGCS
    # ▲ CoxCS
    scale_shape_manual(
      values = c(
        gsgcs = 16,
        coxcs = 17
      ),
      
      breaks = c(
        "gsgcs",
        "coxcs"
      ),
      
      labels = c(
        "GSGCS",
        "CoxCS"
      ),
      
      drop = FALSE,
      
      name = "Method",
      
      guide = guide_legend(
        title.position = "left",
        direction = "horizontal",
        nrow = 1,
        order = 3,
        
        override.aes = list(
          color = "black",
          size  = 4,
          alpha = 1
        )
      )
    ) +
    
    scale_x_discrete(
      limits = hall_pathway_order,
      drop = FALSE,
      expand = expansion(
        add = c(0.55, 0.55)
      )
    ) +
    
    labs(
      title = "Top-ranked enriched pathways",
      x = NULL,
      y = "Cancer type"
    ) +
    
    theme_bw(base_size = 10) +
    
    theme(
      plot.title = element_text(
        hjust = 0.5,
        face = "bold",
        size = 11,
        margin = margin(b = 6)
      ),
      
      axis.text.x = element_text(
        angle = 60,
        hjust = 1,
        vjust = 1,
        size = 7.2
      ),
      
      axis.text.y = element_text(
        size = 8.5
      ),
      
      strip.background = element_rect(
        fill = "grey90",
        color = "grey45",
        linewidth = 0.4
      ),
      
      strip.text.x = element_text(
        face = "bold",
        size = 9
      ),
      
      strip.text.y = element_text(
        face = "bold",
        size = 8.5
      ),
      
      panel.grid.major = element_line(
        color = "grey88",
        linewidth = 0.35
      ),
      
      panel.grid.minor = element_blank(),
      
      panel.spacing.x = unit(
        0.18,
        "cm"
      ),
      
      panel.spacing.y = unit(
        0.30,
        "cm"
      ),
      
      legend.position = "bottom",
      legend.box = "horizontal",
      legend.direction = "horizontal",
      legend.justification = "center",
      
      legend.spacing.x = unit(
        0.40,
        "cm"
      ),
      
      legend.margin = margin(
        t = 5,
        r = 0,
        b = 0,
        l = 0
      ),
      
      plot.margin = margin(
        4,
        5,
        5,
        5
      )
    )
  
  # 13. 最终组合
  # 上 1/3；下 2/3
  p_hallmark <-
    (p_hall_top / p_hall_bottom) +
    plot_layout(
      heights = c(1, 2)
    )
  
  # 14. 显示
  print(p_hallmark)
  
  # 15. 保存
  ggsave(
    filename = "Figure4A_Hallmark_GSGCS_vs_CoxCS_final.png",
    plot = p_hallmark,
    width = 16,
    height = 12,
    units = "in",
    dpi = 300,
    bg = "white"
  )
}
# Figure 4B: KEGG
# 上：GSGCS vs CoxCS 数量 summary strip
# 下：左右分栏 paired bubble plot
#

if(nrow(pdata_kegg) > 0) {
  # Figure 4B / Supplementary Figure
  # KEGG enrichment
  #
  # Main figure:
  #   Top 25 global display pathways
  #
  # Supplementary figure:
  #   Full union of originally selected pathways
  #
  # Global ranking:
  #   1. frequency   : descending
  #   2. max_fold    : descending
  #   3. min_Padj    : ascending
  
  # 0. 基础设置
  cancer_display <- c(
    "BRCA",
    "LUAD",
    "OV",
    "STAD",
    "THYM"
  )
  
  # ggplot 离散 y 轴由下向上排列
  cancer_levels <- rev(cancer_display)
  
  aim_levels <- c(
    "os",
    "pfi",
    "time"
  )
  
  method_levels <- c(
    "gsgcs",
    "coxcs"
  )
  
  aim_labels <- c(
    os   = "OS",
    pfi  = "PFI",
    time = "Time"
  )
  
  method_labels <- c(
    gsgcs = "GSGCS",
    coxcs = "CoxCS"
  )
  
  # 正文最终展示多少个 KEGG pathways
  # 想改成 Top 20 / Top 30，只需要修改这里
  
  KEGG_MAIN_TOP_N <- 25
  
  # PART A
  # Number of significant enriched pathways
  #
  # 注意：
  # 上半部分使用 all_enrich 中全部 Padj < 0.05 pathway
  # 与正文 Top 25 筛选完全无关
  
  # 1. 全部显著 KEGG pathways
  kegg_all <- all_enrich %>%
    
    mutate(
      cancer = toupper(
        trimws(
          as.character(cancer)
        )
      ),
      
      aim = tolower(
        trimws(
          as.character(aim)
        )
      ),
      
      type = tolower(
        trimws(
          as.character(type)
        )
      ),
      
      by = tolower(
        trimws(
          as.character(by)
        )
      ),
      
      kappa_type = trimws(
        as.character(kappa_type)
      )
    ) %>%
    
    mutate(
      type = dplyr::recode(
        type,
        "gsgc"  = "gsgcs",
        "gsgcs" = "gsgcs"
      )
    ) %>%
    
    filter(
      by == "kegg",
      kappa_type == "Theoretical",
      Padj < 0.05
    )
  
  # 检查 method
  bad_methods <- setdiff(
    unique(
      na.omit(
        kegg_all$type
      )
    ),
    method_levels
  )
  
  if(length(bad_methods) > 0) {
    
    stop(
      paste0(
        "KEGG 数据中存在无法识别的方法名称：",
        paste(
          bad_methods,
          collapse = ", "
        ),
        "\n允许值只能是：gsgcs, coxcs"
      )
    )
  }
  
  
  kegg_all <- kegg_all %>%
    
    mutate(
      cancer = factor(
        cancer,
        levels = cancer_levels
      ),
      
      aim = factor(
        aim,
        levels = aim_levels
      ),
      
      type = factor(
        type,
        levels = method_levels
      )
    )
  # 2. 统计全部显著 pathways
  kegg_count <- kegg_all %>%
    
    filter(
      !is.na(cancer),
      !is.na(aim),
      !is.na(type),
      !is.na(Description)
    ) %>%
    
    distinct(
      cancer,
      aim,
      type,
      Description
    ) %>%
    
    count(
      cancer,
      aim,
      type,
      name = "N"
    ) %>%
    
    complete(
      cancer = factor(
        cancer_levels,
        levels = cancer_levels
      ),
      
      aim = factor(
        aim_levels,
        levels = aim_levels
      ),
      
      type = factor(
        method_levels,
        levels = method_levels
      ),
      
      fill = list(
        N = 0
      )
    ) %>%
    
    mutate(
      N_plot = ifelse(
        type == "gsgcs",
        -N,
        N
      )
    )
  # 3. Delta = GSGCS - CoxCS
  kegg_delta <- kegg_count %>%
    
    dplyr::select(
      cancer,
      aim,
      type,
      N
    ) %>%
    
    pivot_wider(
      names_from = type,
      values_from = N,
      values_fill = 0
    ) %>%
    
    mutate(
      delta = gsgcs - coxcs,
      
      delta_label = case_when(
        delta > 0 ~ paste0(
          "+",
          delta
        ),
        
        TRUE ~ as.character(
          delta
        )
      ),
      
      delta_group = case_when(
        delta > 0 ~ "positive",
        delta < 0 ~ "negative",
        TRUE      ~ "zero"
      )
    )
  # 4. 上图范围
  kegg_maxN <- max(
    kegg_count$N,
    na.rm = TRUE
  )
  
  if(
    !is.finite(kegg_maxN) ||
    kegg_maxN <= 0
  ) {
    
    kegg_maxN <- 1
  }
  
  kegg_lim <- kegg_maxN * 1.20
  # 5. mirrored bar plot
  p_kegg_bar <- ggplot(
    kegg_count,
    aes(
      x = N_plot,
      y = cancer,
      fill = type
    )
  ) +
    
    geom_col(
      width = 0.52,
      na.rm = TRUE
    ) +
    
    geom_vline(
      xintercept = 0,
      color = "grey25",
      linewidth = 0.45
    ) +
    
    geom_text(
      aes(
        label = N,
        
        hjust = ifelse(
          type == "gsgcs",
          1.15,
          -0.15
        )
      ),
      
      size = 3,
      color = "grey20",
      na.rm = TRUE
    ) +
    
    facet_grid(
      aim ~ .,
      
      scales = "free_y",
      space = "free_y",
      switch = "y",
      
      labeller = labeller(
        aim = aim_labels
      )
    ) +
    
    scale_fill_manual(
      values = c(
        gsgcs = "#EF6258",
        coxcs = "#4C89CF"
      ),
      
      breaks = c(
        "gsgcs",
        "coxcs"
      ),
      
      labels = c(
        "GSGCS",
        "CoxCS"
      ),
      
      drop = FALSE,
      name = NULL
    ) +
    
    scale_x_continuous(
      limits = c(
        -kegg_lim,
        kegg_lim
      ),
      
      breaks = pretty(
        c(
          -kegg_maxN,
          kegg_maxN
        ),
        n = 5
      ),
      
      labels = function(x) abs(x),
      
      expand = expansion(
        mult = c(
          0.01,
          0.01
        )
      )
    ) +
    
    labs(
      title = "Number of significant enriched pathways",
      x = "Number of pathways",
      y = NULL
    ) +
    
    theme_bw(
      base_size = 10
    ) +
    
    theme(
      plot.title = element_text(
        hjust = 0.5,
        face = "bold",
        size = 11,
        margin = margin(
          b = 5
        )
      ),
      
      axis.text.y = element_text(
        size = 9
      ),
      
      axis.text.x = element_text(
        size = 8
      ),
      
      strip.background = element_rect(
        fill = "grey90",
        color = "grey45",
        linewidth = 0.4
      ),
      
      strip.text.y.left = element_text(
        angle = 0,
        face = "bold",
        size = 8.5
      ),
      
      panel.grid.major.y = element_blank(),
      
      panel.grid.minor = element_blank(),
      
      panel.grid.major.x = element_line(
        color = "grey92",
        linewidth = 0.3
      ),
      
      panel.spacing.y = unit(
        0.45,
        "cm"
      ),
      
      legend.position = "top",
      
      legend.direction = "horizontal",
      
      legend.justification = "center",
      
      legend.key.width = unit(
        0.55,
        "cm"
      ),
      
      legend.key.height = unit(
        0.30,
        "cm"
      ),
      
      plot.margin = margin(
        4,
        2,
        2,
        5
      )
    )
  # 6. Delta
  p_kegg_delta <- ggplot(
    kegg_delta,
    aes(
      x = 1,
      y = cancer
    )
  ) +
    
    geom_text(
      aes(
        label = delta_label,
        color = delta_group
      ),
      
      size = 3.2,
      fontface = "bold"
    ) +
    
    facet_grid(
      aim ~ .,
      
      scales = "free_y",
      space = "free_y"
    ) +
    
    scale_color_manual(
      values = c(
        positive = "#EF4B4B",
        negative = "#377FC1",
        zero = "grey35"
      ),
      
      guide = "none"
    ) +
    
    scale_x_continuous(
      limits = c(
        0.65,
        1.35
      ),
      
      expand = c(
        0,
        0
      )
    ) +
    
    labs(
      title = "\u0394 (GSGCS \u2212 CoxCS)"
    ) +
    
    theme_void(
      base_size = 10
    ) +
    
    theme(
      plot.title = element_text(
        hjust = 0.5,
        face = "bold",
        size = 10,
        margin = margin(
          b = 5
        )
      ),
      
      panel.spacing.y = unit(
        0.45,
        "cm"
      ),
      
      strip.text = element_blank(),
      
      plot.margin = margin(
        4,
        5,
        2,
        0
      )
    )
  # 7. 上半部分组合
  p_kegg_top <-
    (
      p_kegg_bar |
        p_kegg_delta
    ) +
    
    plot_layout(
      widths = c(
        0.94,
        0.06
      )
    )
  
  # PART B
  #
  # 先建立完整 union
  #
  # pdata_kegg 已经是原 cancer × endpoint × method
  # 的 Top-ranked pathway 结果
  
  # 8. 清理完整 bubble 数据
  kegg_bubble_full <- pdata_kegg %>%
    
    mutate(
      cancer = toupper(
        trimws(
          as.character(cancer)
        )
      ),
      
      aim = tolower(
        trimws(
          as.character(aim)
        )
      ),
      
      type = tolower(
        trimws(
          as.character(type)
        )
      )
    ) %>%
    
    mutate(
      type = dplyr::recode(
        type,
        "gsgc"  = "gsgcs",
        "gsgcs" = "gsgcs"
      )
    ) %>%
    
    filter(
      !is.na(cancer),
      !is.na(aim),
      !is.na(type),
      !is.na(Description),
      !is.na(fold),
      !is.na(LOGP),
      is.finite(fold),
      is.finite(LOGP)
    )
  
  
  if(nrow(kegg_bubble_full) == 0) {
    
    stop(
      paste0(
        "kegg_bubble_full 没有可用于绘图的数据。",
        "请检查 pdata_kegg 中的 fold 和 LOGP。"
      )
    )
  }
  # 9. 准备 Padj
  #
  # 优先直接使用 pdata_kegg$Padj
  # 如果没有 Padj，则由 LOGP = -log10(Padj) 反推
  if("Padj" %in% names(kegg_bubble_full)) {
    
    kegg_bubble_full <- kegg_bubble_full %>%
      
      mutate(
        Padj_for_rank = suppressWarnings(
          as.numeric(
            as.character(Padj)
          )
        )
      )
    
  } else {
    
    kegg_bubble_full <- kegg_bubble_full %>%
      
      mutate(
        Padj_for_rank = 10^(-LOGP)
      )
  }
  
  
  # 对异常/缺失 Padj 用 LOGP 反推值补齐
  kegg_bubble_full <- kegg_bubble_full %>%
    
    mutate(
      Padj_for_rank = ifelse(
        is.na(Padj_for_rank) |
          !is.finite(Padj_for_rank),
        
        10^(-LOGP),
        
        Padj_for_rank
      )
    )
  
  # 10. 建立全局 KEGG pathway ranking
  #
  # frequency：
  # pathway 出现在多少个独立的
  # cancer × endpoint × method 组合中
  #
  # 排序：
  # frequency ↓
  # max_fold  ↓
  # min_Padj  ↑
  
  kegg_global_rank <- kegg_bubble_full %>%
    
    group_by(
      Description
    ) %>%
    
    summarise(
      
      frequency = n_distinct(
        interaction(
          cancer,
          aim,
          type,
          drop = TRUE
        )
      ),
      
      max_fold = max(
        fold,
        na.rm = TRUE
      ),
      
      min_Padj = min(
        Padj_for_rank,
        na.rm = TRUE
      ),
      
      max_LOGP = max(
        LOGP,
        na.rm = TRUE
      ),
      
      .groups = "drop"
    ) %>%
    
    arrange(
      desc(frequency),
      desc(max_fold),
      min_Padj,
      Description
    ) %>%
    
    mutate(
      global_rank = row_number()
    )
  # 11. 输出 Supplementary ranking table
  write.csv(
    kegg_global_rank %>%
      
      dplyr::select(
        global_rank,
        Description,
        frequency,
        max_fold,
        min_Padj,
        max_LOGP
      ),
    
    file = "TableS_KEGG_PathwayRanking.csv",
    
    row.names = FALSE
  )
  
  # 12. 正文 Top 25 pathway 集合
  kegg_main_pathways <- kegg_global_rank %>%
    
    slice_head(
      n = KEGG_MAIN_TOP_N
    ) %>%
    
    pull(
      Description
    )
  
  
  # 实际数量
  # 如果总 union < 25，则自然使用全部
  kegg_main_n <- length(
    kegg_main_pathways
  )
  
  
  message(
    "KEGG full union pathways: ",
    nrow(kegg_global_rank)
  )
  
  message(
    "KEGG pathways displayed in main figure: ",
    kegg_main_n
  )
  
  # 13. 正文数据
  kegg_bubble_main <- kegg_bubble_full %>%
    
    filter(
      Description %in% kegg_main_pathways
    )
  
  # 14. pathway 顺序
  #
  # 左 -> 右：
  # global rank 1 -> 25
  
  kegg_main_order <- kegg_global_rank %>%
    
    filter(
      Description %in% kegg_main_pathways
    ) %>%
    
    arrange(
      global_rank
    ) %>%
    
    pull(
      Description
    )
  
  
  kegg_full_order <- kegg_global_rank %>%
    
    arrange(
      global_rank
    ) %>%
    
    pull(
      Description
    )
  
  # 15. factor
  kegg_bubble_main <- kegg_bubble_main %>%
    
    mutate(
      cancer = factor(
        cancer,
        levels = cancer_levels
      ),
      
      aim = factor(
        aim,
        levels = aim_levels
      ),
      
      type = factor(
        type,
        levels = method_levels
      ),
      
      Description = factor(
        Description,
        levels = kegg_main_order
      )
    )
  
  
  kegg_bubble_full <- kegg_bubble_full %>%
    
    mutate(
      cancer = factor(
        cancer,
        levels = cancer_levels
      ),
      
      aim = factor(
        aim,
        levels = aim_levels
      ),
      
      type = factor(
        type,
        levels = method_levels
      ),
      
      Description = factor(
        Description,
        levels = kegg_full_order
      )
    )
  
  # 16. 统一 color / size 范围
  #
  # Main 与 Supplementary 必须共用 scale
  # 才能直接视觉比较
  
  kegg_fold_range <- range(
    kegg_bubble_full$fold,
    na.rm = TRUE
  )
  
  
  if(
    !all(
      is.finite(
        kegg_fold_range
      )
    ) ||
    diff(
      kegg_fold_range
    ) == 0
  ) {
    
    kegg_fold_range <- c(
      min(
        kegg_bubble_full$fold,
        na.rm = TRUE
      ) - 0.01,
      
      max(
        kegg_bubble_full$fold,
        na.rm = TRUE
      ) + 0.01
    )
  }
  
  
  kegg_logp_range <- range(
    kegg_bubble_full$LOGP,
    na.rm = TRUE
  )
  
  
  if(
    !all(
      is.finite(
        kegg_logp_range
      )
    ) ||
    diff(
      kegg_logp_range
    ) == 0
  ) {
    
    kegg_logp_range <- c(
      max(
        0,
        min(
          kegg_bubble_full$LOGP,
          na.rm = TRUE
        ) - 0.01
      ),
      
      max(
        kegg_bubble_full$LOGP,
        na.rm = TRUE
      ) + 0.01
    )
  }
  
  # 17. 通用 bubble plot 函数
  #
  # main / supplementary 共用
  
  make_kegg_bubble_plot <- function(
    plot_data,
    pathway_order,
    title_text,
    x_text_size = 7.2
  ) {
    
    ggplot(
      plot_data,
      aes(
        x = Description,
        y = cancer
      )
    ) +
      
      geom_point(
        aes(
          color = fold,
          size = LOGP,
          shape = type
        ),
        
        alpha = 0.88,
        stroke = 0.6,
        na.rm = TRUE
      ) +
      
      facet_grid(
        rows = vars(
          aim
        ),
        
        cols = vars(
          type
        ),
        
        scales = "fixed",
        space = "fixed",
        
        labeller = labeller(
          aim = aim_labels,
          
          type = c(
            gsgcs = "GSGCS",
            coxcs = "CoxCS"
          )
        )
      ) +
      
      # Fold enrichment
      # reverse Spectral
      
      scale_color_gradientn(
        colours = rev(
          RColorBrewer::brewer.pal(
            11,
            "Spectral"
          )
        ),
        
        limits = kegg_fold_range,
        
        name = "Fold enrichment",
        
        guide = guide_colorbar(
          title.position = "top",
          title.hjust = 0.5,
          
          direction = "horizontal",
          
          barwidth = unit(
            3.5,
            "cm"
          ),
          
          barheight = unit(
            0.30,
            "cm"
          ),
          
          order = 2
        )
      ) +
      
      # -log10(Padj)
      scale_size_continuous(
        limits = kegg_logp_range,
        
        range = c(
          2,
          10
        ),
        
        name = "-log10(Padj)",
        
        guide = guide_legend(
          title.position = "left",
          direction = "horizontal",
          nrow = 1,
          order = 1,
          
          override.aes = list(
            shape = 16,
            color = "black",
            alpha = 1
          )
        )
      ) +
      
      # Method
      # ● GSGCS ▲ CoxCS
      scale_shape_manual(
        values = c(
          gsgcs = 16,
          coxcs = 17
        ),
        
        breaks = c(
          "gsgcs",
          "coxcs"
        ),
        
        labels = c(
          "GSGCS",
          "CoxCS"
        ),
        
        drop = FALSE,
        
        name = "Method",
        
        guide = guide_legend(
          title.position = "left",
          direction = "horizontal",
          nrow = 1,
          order = 3,
          
          override.aes = list(
            color = "black",
            size = 4,
            alpha = 1
          )
        )
      ) +
      # x-axis
      scale_x_discrete(
        limits = pathway_order,
        drop = FALSE,
        
        expand = expansion(
          add = c(
            0.55,
            0.55
          )
        )
      ) +
      
      labs(
        title = title_text,
        x = NULL,
        y = "Cancer type"
      ) +
      
      theme_bw(
        base_size = 10
      ) +
      
      theme(
        plot.title = element_text(
          hjust = 0.5,
          face = "bold",
          size = 11,
          margin = margin(
            b = 6
          )
        ),
        
        axis.text.x = element_text(
          angle = 60,
          hjust = 1,
          vjust = 1,
          size = x_text_size
        ),
        
        axis.text.y = element_text(
          size = 8.5
        ),
        
        strip.background = element_rect(
          fill = "grey90",
          color = "grey45",
          linewidth = 0.4
        ),
        
        strip.text.x = element_text(
          face = "bold",
          size = 9
        ),
        
        strip.text.y = element_text(
          face = "bold",
          size = 8.5
        ),
        
        panel.grid.major = element_line(
          color = "grey88",
          linewidth = 0.35
        ),
        
        panel.grid.minor = element_blank(),
        
        panel.spacing.x = unit(
          0.18,
          "cm"
        ),
        
        panel.spacing.y = unit(
          0.30,
          "cm"
        ),
        
        legend.position = "bottom",
        
        legend.box = "horizontal",
        
        legend.direction = "horizontal",
        
        legend.justification = "center",
        
        legend.spacing.x = unit(
          0.40,
          "cm"
        ),
        
        legend.margin = margin(
          t = 5,
          r = 0,
          b = 0,
          l = 0
        ),
        
        plot.margin = margin(
          4,
          5,
          5,
          5
        )
      )
  }
  
  # PART C
  # Main Figure 4B
  #
  # 只显示 global Top 25
  
  # 18. Main bubble plot
  
  p_kegg_bottom_main <- make_kegg_bubble_plot(
    plot_data = kegg_bubble_main,
    
    pathway_order = kegg_main_order,
    
    title_text = "Global Top 25 enriched pathways",
    
    # 25 个 pathway 后可适当增大字体
    x_text_size = 8.0
  )
  
  # 19. Main Figure 4B
  p_kegg_main <-
    (
      p_kegg_top /
        p_kegg_bottom_main
    ) +
    
    plot_layout(
      heights = c(
        1,
        2
      )
    )
  
  # 20. 显示 Main Figure
  print(
    p_kegg_main
  )
  
  # 21. 保存 Main Figure
  ggsave(
    filename = "Figure4B_KEGG_Main_Top25.png",
    
    plot = p_kegg_main,
    
    width = 16,
    height = 12,
    
    units = "in",
    
    dpi = 300,
    
    bg = "white"
  )
  
  # PART D
  # Supplementary Figure
  #
  # 完整 union
  
  # 22. Supplementary bubble plot
  p_kegg_bottom_full <- make_kegg_bubble_plot(
    plot_data = kegg_bubble_full,
    
    pathway_order = kegg_full_order,
    
    title_text = "Full union of top-ranked enriched KEGG pathways",
    
    # pathway 多，因此字体稍小
    x_text_size = 6.2
  )
  
  # 23. Supplementary Figure
  #
  # Supplementary 没必要再重复上方 count panel。
  # 它的任务是完整展示 enrichment profile。
  p_kegg_supp <- p_kegg_bottom_full
  
  # 24. 显示 Supplementary Figure
  print(
    p_kegg_supp
  )
  
  
  # 25. 保存 Supplementary Figure
  #
  # 完整 union 横轴很多，
  # 因此宽度比正文更大
  ggsave(
    filename = "FigureS_KEGG_FullUnion.png",
    
    plot = p_kegg_supp,
    
    width = 20,
    height = 9,
    
    units = "in",
    
    dpi = 300,
    
    bg = "white"
  )
  
  # 26. 同时输出正文实际展示的 25 个 pathway
  kegg_main_table <- kegg_global_rank %>%
    
    filter(
      global_rank <= KEGG_MAIN_TOP_N
    ) %>%
    
    dplyr::select(
      global_rank,
      Description,
      frequency,
      max_fold,
      min_Padj,
      max_LOGP
    )
  
  
  write.csv(
    kegg_main_table,
    
    file = "Figure4B_KEGG_Main_Top25_Pathways.csv",
    
    row.names = FALSE
  )
  
  # 27. 控制台检查
  cat(
    "\n========================================\n"
  )
  
  cat(
    "KEGG Figure generation completed\n"
  )
  
  cat(
    "Full union pathways:",
    nrow(kegg_global_rank),
    "\n"
  )
  
  cat(
    "Main figure pathways:",
    kegg_main_n,
    "\n"
  )
  
  cat(
    "Main figure:",
    "Figure4B_KEGG_Main_Top25.png\n"
  )
  
  cat(
    "Supplementary figure:",
    "FigureS_KEGG_FullUnion.png\n"
  )
  
  cat(
    "Full ranking table:",
    "TableS_KEGG_PathwayRanking.csv\n"
  )
  
  cat(
    "Main pathway table:",
    "Figure4B_KEGG_Main_Top25_Pathways.csv\n"
  )
  
  cat(
    "========================================\n\n"
  )
}




#### （3.5）富集结果绘制 ： Figure 4C ----
# 1. Parameters
cancer_order   <- c("BRCA", "LUAD", "OV", "STAD", "THYM")
method_order   <- c("coxcs", "gsgcs")
endpoint_order <- c("os", "pfi", "time")
output_dir     <- "Figure4C_final"
if (!dir.exists(output_dir)) { dir.create(output_dir, recursive = TRUE) }

# 2. Check input
if (!exists("sig_kegg")) { 
  stop("Object 'sig_kegg' does not exist.") }
required_cols <- c("cancer", "type", "aim", "Description")
missing_cols  <- setdiff(required_cols, names(sig_kegg))
if (length(missing_cols) > 0) { 
  stop(paste0("Missing columns in sig_kegg: ", paste(missing_cols, collapse = ", "))) }

# 3. Clean data
sig_kegg_C <- sig_kegg %>% 
  mutate(cancer = toupper(trimws(as.character(cancer))), 
         type = tolower(trimws(as.character(type))), 
         aim = tolower(trimws(as.character(aim))), 
         Description = as.character(Description)) %>% 
  mutate(type = case_when(type %in% c("gsgc", "gsgcs") ~ "gsgcs", 
                          type == "coxcs" ~ "coxcs", TRUE ~ type)) %>% 
  filter(cancer %in% cancer_order, 
         type %in% method_order, 
         aim %in% endpoint_order, 
         !is.na(Description), 
         Description != "") %>% 
  distinct(cancer, type, aim, Description)

# 4. Check cleaned data
cat("\n========== Data availability ==========\n")
availability_check <- sig_kegg_C %>% 
  count(cancer, type, aim, name = "n_pathways") %>% 
  arrange(cancer, type, aim)
print(availability_check, n = Inf)

# 5. Safe Jaccard
jaccard_safe <- function(A, B) { 
  A <- unique(A[!is.na(A)]); 
  B <- unique(B[!is.na(B)]); 
  if (length(A) == 0 || length(B) == 0) { 
    return(NA_real_) }; 
  union_AB <- union(A, B); 
  if (length(union_AB) == 0) { 
    return(NA_real_) }; 
  length(intersect(A, B)) / length(union_AB) }

# 6. Summary function
calculate_summary <- function(cancer_name, method_name) { 
  dat <- sig_kegg_C %>%  filter(cancer == cancer_name, type == method_name); 
  OS_set   <- dat %>% filter(aim == "os") %>% pull(Description) %>% unique(); 
  PFI_set  <- dat %>% filter(aim == "pfi") %>% pull(Description) %>% unique(); 
  Time_set <- dat %>% filter(aim == "time") %>% pull(Description) %>% unique(); 
  J_OS_PFI <- jaccard_safe(OS_set, PFI_set); 
  tibble(cancer = cancer_name, 
         type = method_name, 
         n_OS = length(OS_set), 
         n_PFI = length(PFI_set), 
         n_Time = length(Time_set), 
         OS_PFI_intersection = length(intersect(OS_set, PFI_set)), 
         OS_PFI_union = length(union(OS_set, PFI_set)), 
         Jaccard_OS_PFI = J_OS_PFI, 
         evaluable_OS_PFI = !is.na(J_OS_PFI)) }

# 7. Calculate all cancer × method combinations
summary_C <- expand_grid(cancer = cancer_order, type = method_order) %>% 
  pmap_dfr(function(cancer, type) { calculate_summary(cancer_name = cancer, method_name = type) })

# 8. Add factors / labels
summary_C <- summary_C %>% 
  mutate(cancer = factor(cancer, levels = cancer_order), 
         type = factor(type, levels = method_order), 
         method_display = case_when(as.character(type) == "coxcs" ~ "CoxCS", 
                                    as.character(type) == "gsgcs" ~ "GSGCS", 
                                    TRUE ~ as.character(type))) %>% 
  arrange(cancer, type)

# 9. Print summary
cat("\n========== Figure 4C summary ==========\n")
print(summary_C, n = Inf, width = Inf)

# 10. Export data
write.csv(summary_C, file = file.path(output_dir, "Figure4C_summary_data.csv"), row.names = FALSE)

# 11. Colors
method_fill   <- c("coxcs" = "#9ECAE1", "gsgcs" = "#FC9272")
method_border <- c("coxcs" = "#2171B5", "gsgcs" = "#DE2D26")

# 12. Prepare Jaccard data
jaccard_dat <- summary_C %>% 
  mutate(cancer_num = as.numeric(cancer), 
         x_position = case_when(as.character(type) == "coxcs" ~ cancer_num - 0.16, 
                                as.character(type) == "gsgcs" ~ cancer_num + 0.16))

# 13. Paired-line data
line_dat <- jaccard_dat %>% group_by(cancer) %>% filter(n() == 2, all(evaluable_OS_PFI)) %>% ungroup()

# 14. Evaluable points
jaccard_eval <- jaccard_dat %>% filter(evaluable_OS_PFI)

# 15. Non-evaluable points
jaccard_NE <- jaccard_dat %>% filter(!evaluable_OS_PFI) %>% mutate(plot_y = 0.035)

# 16. Upper panel
p_upper <- ggplot()
p_upper <- p_upper + 
  geom_vline(xintercept = c(1.5, 2.5, 3.5, 4.5), 
             colour = "grey88", linewidth = 0.45, linetype = "dashed")
p_upper <- p_upper + 
  geom_line(data = line_dat, aes(x = x_position, y = Jaccard_OS_PFI, group = cancer), 
            colour = "grey35", linewidth = 0.85)
p_upper <- p_upper + 
  geom_point(data = jaccard_eval, aes(x = x_position, y = Jaccard_OS_PFI, fill = type, colour = type), 
             shape = 21, size = 5.3, stroke = 1.2)
p_upper <- p_upper + 
  geom_text(data = jaccard_eval, aes(x = x_position, y = Jaccard_OS_PFI, label = sprintf("%.3f", Jaccard_OS_PFI), colour = type), 
            vjust = -1.05, size = 4.0, fontface = "bold", show.legend = FALSE)
if (nrow(jaccard_NE) > 0) { 
  p_upper <- p_upper + 
    geom_point(data = jaccard_NE, aes(x = x_position, y = plot_y, colour = type), 
               shape = 1, size = 5, stroke = 1.2, show.legend = FALSE) + 
    geom_text(data = jaccard_NE, aes(x = x_position, y = plot_y, label = "NE"), 
              vjust = -1.15, size = 3.8, fontface = "bold", colour = "grey25") }
p_upper <- p_upper + 
  scale_fill_manual(  values = method_fill,   breaks = c("coxcs", "gsgcs"), labels = c("CoxCS", "GSGCS")) + 
  scale_colour_manual(values = method_border, breaks = c("coxcs", "gsgcs"), labels = c("CoxCS", "GSGCS")) + 
  scale_x_continuous(breaks = seq_along(cancer_order), labels = cancer_order, limits = c(0.55, 5.45), expand = c(0, 0)) + 
  scale_y_continuous(limits = c(0, 1.05), breaks = c(0, 0.25, 0.50, 0.75, 1.00), labels = c("0.00", "0.25", "0.50", "0.75", "1.00"), expand = expansion(mult = c(0, 0))) + 
  labs(title = "Cross-endpoint pathway recurrence (OS–PFI) across cancers", 
       subtitle = "Jaccard similarity of significantly enriched KEGG pathways", 
       x = NULL, y = "Jaccard similarity\n(OS vs PFI)") + 
  theme_classic(base_size = 12) + 
  theme(plot.title = element_text(size = 16, face = "bold", hjust = 0.5, margin = margin(b = 3)), 
        plot.subtitle = element_text(size = 10.5, colour = "grey40", hjust = 0.5, margin = margin(b = 10)), 
        axis.title.y = element_text(size = 11.5, face = "bold", margin = margin(r = 8)), 
        axis.text.y = element_text(size = 10, colour = "black"), 
        axis.text.x = element_text(size = 11, face = "bold", colour = "black", margin = margin(t = 5)), 
        axis.ticks.x = element_blank(), legend.position = "none", plot.margin = margin(8, 10, 8, 8))

# 17. Convert to long format
count_dat <- summary_C %>% 
  select(cancer, type, method_display, n_OS, n_PFI, n_Time) %>% 
  pivot_longer(cols = c(n_OS, n_PFI, n_Time), names_to = "endpoint", values_to = "n_pathways") %>% 
  mutate(endpoint = recode(endpoint, "n_OS" = "OS", "n_PFI" = "PFI", "n_Time" = "PCAWG"), 
         endpoint = factor(endpoint, levels = c("OS", "PFI", "PCAWG")), 
         cancer = factor(cancer, levels = cancer_order), 
         type = factor(type, levels = method_order))

# 18. Automatic y-axis
max_count <- max(count_dat$n_pathways, na.rm = TRUE)
y_upper   <- ceiling(max_count / 50) * 50 + 25

# 19. Lower panel
p_lower <- ggplot(count_dat, aes(x = endpoint, y = n_pathways, fill = type, colour = type)) + 
  geom_col(position = position_dodge(width = 0.76), width = 0.66, linewidth = 0.65) + 
  geom_text(aes(label = n_pathways), 
            position = position_dodge(width = 0.76), 
            vjust = -0.45, size = 3.0, show.legend = FALSE) + 
  facet_grid(. ~ cancer, scales = "fixed", switch = "x") + 
  scale_fill_manual(values = method_fill, breaks = c("coxcs", "gsgcs"), labels = c("CoxCS", "GSGCS")) + 
  scale_colour_manual(values = method_border, breaks = c("coxcs", "gsgcs"), labels = c("CoxCS", "GSGCS")) + 
  scale_y_continuous(limits = c(0, y_upper), breaks = seq(0, y_upper, by = 100), expand = expansion(mult = c(0, 0))) + 
  labs(title = "Number of significantly enriched KEGG pathways by endpoint", 
       subtitle = "Adjusted P < 0.05", 
       x = NULL, y = "Number of pathways", fill = NULL, colour = NULL) + 
  theme_classic(base_size = 11) + 
  theme(plot.title = element_text(size = 14, face = "bold", hjust = 0.5, margin = margin(b = 2)), 
        plot.subtitle = element_text(size = 9.5, colour = "grey40", hjust = 0.5, margin = margin(b = 8)), 
        axis.title.y = element_text(size = 11, face = "bold", margin = margin(r = 7)), 
        axis.text.y = element_text(size = 9.5, colour = "black"), 
        axis.text.x = element_text(size = 9.5, face = "bold", colour = "black"), 
        axis.ticks.x = element_blank(), strip.background = element_blank(), 
        strip.text.x = element_text(size = 11, face = "bold", colour = "black", margin = margin(t = 6, b = 4)), 
        panel.spacing.x = unit(0.45, "cm"), legend.position = "right", legend.text = element_text(size = 10), 
        legend.title = element_blank(), legend.key.height = unit(0.5, "cm"), plot.margin = margin(5, 10, 8, 8))

# 20. Combine panels
figure_body <- p_upper / p_lower + plot_layout(heights = c(1.15, 1))
panel_C  <- wrap_elements(full = textGrob("C", 
                                          x = unit(0.02, "npc"), y = unit(0.50, "npc"), 
                                          just = c("left", "center"), 
                                          gp = gpar(fontsize = 22, fontface = "bold")))
Figure4C <- panel_C | figure_body
Figure4C <- Figure4C + plot_layout(widths = c(0.045, 1))

# 21. Display
print(Figure4C)

# 22. Save full Figure 4C — PNG
ggsave(filename = file.path(output_dir, "Figure4C_cross_endpoint_recurrence.png"), 
       plot = Figure4C, width = 13, height = 9.5, units = "in", dpi = 600, bg = "white", limitsize = FALSE)

# 23. Save full Figure 4C — PDF
ggsave(filename = file.path(output_dir, "Figure4C_cross_endpoint_recurrence.pdf"), 
       plot = Figure4C, width = 13, height = 9.5, units = "in", device = cairo_pdf, bg = "white", limitsize = FALSE)

# 24. Save upper panel separately
ggsave(filename = file.path(output_dir, "Figure4C_OS_PFI_Jaccard.png"), 
       plot = p_upper, width = 11, height = 5.2, units = "in", dpi = 600, bg = "white")

# 25. Save lower panel separately
ggsave(filename = file.path(output_dir, "Figure4C_pathway_counts.png"), 
       plot = p_lower, width = 13, height = 4.5, units = "in", dpi = 600, bg = "white")

# 26. Final numerical check
cat("\n\n========== Final numerical check ==========\n")
check_table <- summary_C %>% 
  select(cancer, type, method_display, 
         n_OS, n_PFI, n_Time, OS_PFI_intersection, OS_PFI_union, Jaccard_OS_PFI, evaluable_OS_PFI) %>% 
  arrange(cancer, type)
print(check_table, n = Inf, width = Inf)

# 27. Expected results
cat("\nExpected OS–PFI Jaccard pattern:\n\n", "BRCA : CoxCS ~0.682 | GSGCS ~0.909\n", "LUAD : CoxCS ~0.525 | GSGCS ~0.865\n", "OV   : CoxCS NE     | GSGCS ~0.907\n", "STAD : CoxCS ~0.489 | GSGCS ~0.926\n", "THYM : CoxCS ~0.718 | GSGCS ~0.751\n\n", "NE = not evaluable because one endpoint contains no ", "significantly enriched pathways.\n")


#### （3.5）富集结果绘制 ： figure 4 D ----
# Recurrent functional organization
# One figure per cancer
#
# BRCA / LUAD / OV / STAD / THYM
#
# Upper:
#   CoxCS union network | GSGCS union network
#
# Middle:
#   common legend
#
# Lower:
#   CoxCS functional-cluster recurrence table
#   GSGCS functional-cluster recurrence table
#
# KEGG + theoretical + Padj < 0.05
#
# Endpoint:
#   OS   = TCGA OS
#   PFI  = TCGA PFI
#   time = PCAWG

# 0. Packages
library(dplyr)
library(tidyr)
library(stringr)
library(purrr)
library(ggplot2)
library(ggrepel)
library(igraph)
library(patchwork)
library(scales)

# 1. Parameters
cancer_list <- c("BRCA","LUAD","OV","STAD","THYM")
endpoint_order <- c("os","pfi","time")
padj_cutoff <- 0.05

# pathway similarity threshold
jaccard_cutoff <- 0.20

# network labels
max_pathway_labels <- 8

# bottom table
top_n_clusters <- 6


# IMPORTANT:
# functional cluster column
#
# Recommended:
# "Level_2_Category"
#
# If you want your own custom category:
# cluster_col <- "小类"
cluster_col <- "Level_2_Category"

# Figure size
figure_width  <- 14
figure_height <- 11

# Output directory
output_dir <- "Figure4_functional_consistency"
if (!dir.exists(output_dir)) {
  dir.create(output_dir,recursive = TRUE)
}


# 2. Check input object
#
# Assumes your enrichment data are stored in:
# df
if (!exists("df")) {
  stop("Object 'df' does not exist. Please load your enrichment data first.")
}

# 3. Clean data
df2 <- df %>%
  mutate(
    cancer = toupper(as.character(cancer)),
    type = tolower(as.character(type)),
    by = tolower(as.character(by)),
    aim = tolower(as.character(aim)),
    kappa_type = tolower(as.character(kappa_type))
  ) %>%
  mutate(type =case_when(type %in% c("gsgc","gsgcs") ~"gsgcs",type == "coxcs" ~"coxcs",TRUE ~type))

# 4. Check required columns
required_columns <- c("cancer","type","by","aim","kappa_type",
                      "ID","Description",
                      "Padj","Count","geneID",
                      "Level_1_Category",cluster_col)

missing_columns <- setdiff(required_columns,names(df2))

if (length(missing_columns) > 0) {
  stop(paste0("Missing required columns: ",paste(missing_columns,collapse = ", ")))
}

# 5. Significant KEGG data
sig_kegg <- df2 %>%
  filter(cancer %in%cancer_list,
         by =="kegg",kappa_type =="theoretical",
         !is.na(Padj),Padj < padj_cutoff,
         type %in% c("coxcs","gsgcs"),
         aim %in% endpoint_order,!is.na(Description),Description != "") %>%
  distinct(cancer,type,aim,Description,.keep_all =TRUE)

cat("\nSignificant KEGG rows:", nrow(sig_kegg),"\n")

# 6. Helper:
# safely return first non-missing value
first_non_missing <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) == 0) {return(NA_character_)}
  as.character( x[1])
}

# 7. Helper:
# safely return first non-empty value
first_non_empty <- function(x) {
  x <- as.character(x)
  x <- x[!is.na(x) &x != ""]
  if (length(x) == 0) {return(NA_character_) }
  x[1]
}

# 8. Split gene list
split_genes <- function(x) {
  if (length(x) == 0 ||is.na(x) ||x == "") {return(character(0))
  }
  x <- str_replace_all(as.character(x),"\\s+","" )
  genes <- unlist(str_split(x,"/|;|,"))
  genes <- genes[genes != ""]
  unique(genes)
}

# 9. Gene-set Jaccard
gene_jaccard <- function(x,y) {
  gx <- split_genes(x)
  gy <- split_genes(y)
  if (length(gx) == 0 ||length(gy) == 0) {return(0)}
  union_genes <- union( gx,gy)
  if (length(union_genes) == 0) { return(0) }
  length(intersect(gx, gy)) /length(union_genes)
}

# 10. Three-way Jaccard
# Important:
# if any endpoint is empty -> NA
three_way_jaccard <- function( os_set,pfi_set, time_set) {
  if (length(os_set) == 0 ||length(pfi_set) == 0 ||length(time_set) == 0) {return(NA_real_) }
  intersection_all <- Reduce(intersect,list(os_set,pfi_set,time_set))
  union_all <- Reduce(union,list(os_set,pfi_set,time_set))
  if (length(union_all) == 0) {return(NA_real_)}
  length(intersection_all) /length(union_all)
}

# 11. Recurrence annotation
# One pathway:
# 1 endpoint
# 2 endpoints
# 3 endpoints
make_recurrence_df <- function(cancer_name,method_name) {
  dat <- sig_kegg %>% filter( cancer ==cancer_name,type == method_name)
  if (nrow(dat) == 0) {return(tibble())}
  dat %>%
    distinct( Description,aim) %>%
    group_by(Description) %>%
    summarise(in_OS = any(aim == "os"),in_PFI = any( aim == "pfi"),in_Time =any(aim == "time"),n_endpoints = n_distinct(aim),.groups ="drop") %>%
    mutate(recurrence_class =case_when( n_endpoints >= 3 ~"Present in all 3 endpoints",
                                        n_endpoints == 2 ~"Present in 2 endpoints",
                                        TRUE ~"Present in 1 endpoint"))
}

# 12. Collapse pathway across endpoints
# ★ This is the section that caused your error.
# Correct dynamic-column syntax
# .data[[cluster_col]]
# NOT:.data[[cluster_col]]
collapse_pathway_nodes <- function(cancer_name,method_name) {
  dat <- sig_kegg %>%filter(cancer ==cancer_name,type == method_name)
  if (nrow(dat) == 0) {return(tibble())}
  recurrence_df <-make_recurrence_df( cancer_name,method_name)
  
  # One pathway = one node
  nodes <- dat %>%arrange(Padj) %>%group_by(Description) %>%
    
    summarise(
      ID = first_non_empty(ID),
      geneID =first_non_empty(geneID),
      Count = if (all(is.na(Count)))  {NA_real_}  else {max(Count,na.rm = TRUE)},
      Padj  = if (all(is.na(Padj))) {NA_real_ } else {min(Padj,na.rm = TRUE)},
      Level_1_Category =first_non_empty(Level_1_Category),
      # CORRECTED LINE
      Functional_cluster = first_non_empty(.data[[cluster_col]]),.groups ="drop")
  
  # Recurrence
  nodes <- nodes %>%left_join(recurrence_df,by ="Description") %>%
    mutate(
      n_endpoints =replace_na(n_endpoints,1L ),
      recurrence_class =replace_na(recurrence_class,"Present in 1 endpoint"),
      in_OS =replace_na(in_OS,FALSE),
      in_PFI =replace_na( in_PFI,FALSE ),
      in_Time =replace_na(in_Time,FALSE))
  nodes
}

# 13. Fixed KEGG Level-1 colors
canonical_colors <- c("Cellular Processes" ="#F8766D",
                      "Environmental Information Processing" ="#C49A00",
                      "Genetic Information Processing" ="#00BA38",
                      "Human Diseases" ="#00BFC4",
                      "Metabolism" ="#619CFF",
                      "Organismal Systems" ="#F564E3")

# Determine actual categories
level1_levels <- sig_kegg %>%
  filter(!is.na(Level_1_Category)) %>%
  pull(Level_1_Category) %>%
  as.character() %>%unique() %>%sort()

unknown_levels <- setdiff(level1_levels,names(canonical_colors))
if (length(unknown_levels) > 0) {
  extra_colors <- setNames(hue_pal()(length(unknown_levels)),unknown_levels)
  category_colors <- c(canonical_colors[intersect(names(canonical_colors),level1_levels) ],extra_colors)
} else {
  category_colors <-canonical_colors[level1_levels]
}

# 14. Build union network
build_union_network <- function(cancer_name, method_name, cutoff = jaccard_cutoff) {
  nodes <- collapse_pathway_nodes(cancer_name, method_name)
  if (nrow(nodes) == 0) {
    return(list(nodes = tibble(), edges = tibble()))
  }
  if (nrow(nodes) >= 2) {
    combinations <- combn(seq_len(nrow(nodes)), 2)
    edge_list <- vector("list", ncol(combinations))
    for (i in seq_len(ncol(combinations))) {
      a <- combinations[1, i]
      b <- combinations[2, i]
      jac <- gene_jaccard(nodes$geneID[a], nodes$geneID[b])
      if (!is.na(jac) && jac >= cutoff) {
        edge_list[[i]] <- tibble(from = nodes$Description[a], to = nodes$Description[b], jaccard = jac)
      }
    }
    edges <- bind_rows(edge_list)
  } else {
    edges <- tibble()
  }
  if (nrow(edges) > 0) {
    graph_object <- graph_from_data_frame(edges %>% select(from, to), directed = FALSE, vertices = data.frame(name = nodes$Description))
    set.seed(2026)
    layout_matrix <- layout_with_fr(graph_object, niter = 2000)
    layout_df <- tibble(Description = V(graph_object)$name, x = layout_matrix[, 1], y = layout_matrix[, 2])
    if (ecount(graph_object) > 0) {
      community_object <- cluster_louvain(graph_object)
      community_df <- tibble(Description = names(membership(community_object)), network_cluster = as.integer(membership(community_object)))
    } else {
      community_df <- tibble(Description = nodes$Description, network_cluster = seq_len(nrow(nodes)))
    }
  } else {
    n_nodes <- nrow(nodes)
    theta <- seq(0, 2 * pi, length.out = n_nodes + 1)[-(n_nodes + 1)]
    layout_df <- tibble(Description = nodes$Description, x = cos(theta), y = sin(theta))
    community_df <- tibble(Description = nodes$Description, network_cluster = seq_len(n_nodes))
  }
  nodes <- nodes %>% left_join(layout_df, by = "Description") %>% left_join(community_df, by = "Description")
  if (nrow(edges) > 0) {
    edges_plot <- edges %>% left_join(layout_df %>% transmute(from = Description, x = x, y = y), by = "from") %>% left_join(layout_df %>% transmute(to = Description, xend = x, yend = y), by = "to")
  } else {
    edges_plot <- tibble(from = character(), to = character(), jaccard = numeric(), x = numeric(), y = numeric(), xend = numeric(), yend = numeric())
  }
  list(nodes = nodes, edges = edges_plot)
}


# 15. Method summary
get_method_summary <- function(cancer_name, method_name) {
  dat <- sig_kegg %>% filter(cancer == cancer_name, type == method_name)
  os_set <- dat %>% filter(aim == "os") %>% pull(Description) %>% unique()
  pfi_set <- dat %>% filter(aim == "pfi") %>% pull(Description) %>% unique()
  time_set <- dat %>% filter(aim == "time") %>% pull(Description) %>% unique()
  if (length(os_set) > 0 && length(pfi_set) > 0 && length(time_set) > 0) {
    all3 <- Reduce(intersect, list(os_set, pfi_set, time_set))
  } else {
    all3 <- character(0)
  }
  j3 <- three_way_jaccard(os_set, pfi_set, time_set)
  tibble(cancer = cancer_name, type = method_name, n_OS = length(os_set), n_PFI = length(pfi_set), n_Time = length(time_set), shared_all3 = length(all3), Jaccard_all3 = j3, evaluable_all3 = !is.na(j3))
}

# 16. Representative labels
select_network_labels <- function(nodes, max_labels = max_pathway_labels) {
  if (nrow(nodes) == 0) {
    return(nodes)
  }
  nodes %>% filter(!is.na(Level_1_Category)) %>% arrange(desc(n_endpoints), Padj) %>% group_by(Level_1_Category) %>% slice_head(n = 2) %>% ungroup() %>% arrange(desc(n_endpoints), Padj) %>% slice_head(n = max_labels)
}


# 17. Plot union network
plot_union_network <- function(network,cancer_name,method_name,count_limits,count_breaks) {
  nodes <-network$nodes
  edges <-network$edges
  method_display <-ifelse(method_name =="gsgcs","GSGCS","CoxCS")
  # Empty
  if (nrow(nodes) == 0) {
    return(
      ggplot() +
        annotate("text",x =0.5,y =0.5,
                 label ="No significantly enriched\nKEGG pathways",
                 size =5,colour ="grey50",fontface ="bold") +xlim(0,1) +ylim(0,1) +
        labs(title =paste0(method_display," | ",cancer_name)) +theme_void())
  }
  # Summary
  summary_row <-get_method_summary(cancer_name,method_name)
  if (isTRUE(summary_row$evaluable_all3)) {
    subtitle_text <-paste0("Shared across all three: ",summary_row$shared_all3,"  |  Three-way Jaccard = ",sprintf("%.2f",summary_row$Jaccard_all3))
  } else {
    subtitle_text <-paste0("Shared across all three: ",summary_row$shared_all3, "  |  Three-way Jaccard: not evaluable")
  }
  nodes_1 <- nodes %>%filter(n_endpoints == 1)
  nodes_2 <- nodes %>%filter(n_endpoints == 2)
  nodes_3 <- nodes %>%filter(n_endpoints >= 3)
  label_nodes <-select_network_labels(nodes)
  # Plot
  p <- ggplot()
  # Edges
  if (nrow(edges) > 0) {
    p <- p +geom_segment(data = edges,
                         aes(x =x,y =y,xend = xend,yend = yend),
                         colour = "grey72",linewidth = 0.25,alpha =0.25)
  }
  # Endpoint-specific
  if (nrow(nodes_1) > 0) {
    p <- p + geom_point(data =nodes_1,
                        aes(x =x,y =y,size =Count,fill = Level_1_Category),
                        shape =21,colour ="grey75",stroke = 0.25,alpha =0.25)
  }
  # Two endpoints
  if (nrow(nodes_2) > 0) {
    p <- p +geom_point(data =nodes_2, 
                       aes( x =x, y =y,size =Count,fill =Level_1_Category ),
                       shape =21,colour ="grey30",stroke =0.75, alpha =0.82)
  }
  # Three endpoints
  if (nrow(nodes_3) > 0) {
    p <- p +geom_point(data =nodes_3,
                       aes(x =x,y =y,size =Count,fill = Level_1_Category),
                       shape =21,colour ="black", stroke =1.15,alpha =1)
  }
  
  # Labels
  if (nrow(label_nodes) > 0) {
    p <- p +geom_text_repel(data =label_nodes,
                            aes( x =x,y =y,label =Description),
                            size =3,colour ="grey10", box.padding = 0.35,
                            point.padding =0.25,
                            min.segment.length =0,segment.size =0.25,segment.color ="grey55", 
                            max.overlaps =Inf, seed =2026)
  }
  
  # Scales
  p <- p +
    scale_fill_manual(values =category_colors,name ="KEGG functional category",drop =FALSE,na.value ="grey80") +
    scale_size_continuous(name ="Prioritized genes\nin pathway",range =c( 2.5,8),limits =count_limits,breaks =count_breaks) +
    labs(title =paste0(method_display," | ",cancer_name),subtitle =subtitle_text) +
    coord_equal( clip ="off") +
    theme_void() +
    theme(
      plot.title = element_text(size =15,face ="bold", hjust =0.5),
      plot.subtitle = element_text(size =9,hjust =0.5,colour ="grey30",margin = margin(b = 6 )),
      legend.position ="none",
      plot.margin = margin(8,10,8, 10))
  p
}

# 18. Functional-cluster summary
#
# Correct dynamic-column syntax again:
#
# .data[[cluster_col]]
make_cluster_summary <- function(cancer_name, method_name) {
  dat <- sig_kegg %>% filter(cancer == cancer_name, type == method_name) %>% mutate(Functional_cluster = as.character(.data[[cluster_col]])) %>% filter(!is.na(Functional_cluster), Functional_cluster != "")
  if (nrow(dat) == 0) { return(tibble()) }
  pathway_endpoint <- dat %>% distinct(Functional_cluster, Description, aim)
  endpoint_counts <- pathway_endpoint %>% count(Functional_cluster, aim, name = "n") %>% complete(Functional_cluster, aim = endpoint_order, fill = list(n = 0)) %>% pivot_wider(names_from = aim, values_from = n, values_fill = 0)
  if (!"os" %in% names(endpoint_counts)) { endpoint_counts$os <- 0 }
  if (!"pfi" %in% names(endpoint_counts)) { endpoint_counts$pfi <- 0 }
  if (!"time" %in% names(endpoint_counts)) { endpoint_counts$time <- 0 }
  pathway_recurrence <- pathway_endpoint %>% group_by(Functional_cluster, Description) %>% summarise(n_endpoints = n_distinct(aim), .groups = "drop")
  recurrence_counts <- pathway_recurrence %>% group_by(Functional_cluster) %>% summarise(shared_all3 = sum(n_endpoints >= 3), shared_2plus = sum(n_endpoints >= 2), total_unique = n_distinct(Description), .groups = "drop")
  endpoint_counts %>% left_join(recurrence_counts, by = "Functional_cluster") %>% mutate(cancer = cancer_name, type = method_name) %>% select(cancer, type, Functional_cluster, os, pfi, time, shared_all3, shared_2plus, total_unique) %>% arrange(desc(shared_all3), desc(shared_2plus), desc(total_unique))
}

# 19. Select Top clusters
select_top_clusters <- function(cluster_summary, n = top_n_clusters) {
  if (nrow(cluster_summary) == 0) {
    return(cluster_summary)
  }
  cluster_summary %>%
    arrange(desc(shared_all3), desc(shared_2plus), desc(total_unique)) %>%
    slice_head(n = n)
}
# 20. Plot cluster table
plot_cluster_table <- function(cluster_summary, cancer_name, method_name) {
  method_display <- ifelse(method_name == "gsgcs", "GSGCS", "CoxCS")
  top <- select_top_clusters(cluster_summary)
  # Empty
  if (nrow(top) == 0) {
    return(ggplot() + annotate("text", x = 0.5, y = 0.5, label = "No functional clusters", size = 4, colour = "grey50") + xlim(0, 1) + ylim(0, 1) + labs(title = paste0(method_display, " | ", cancer_name)) + theme_void())
  }
  top <- top %>% mutate(row = rev(seq_len(n())))
  n_rows <- nrow(top)
  p <- ggplot() + xlim(0, 10) + ylim(0, n_rows + 2.40)
  # Header color
  header_fill <- ifelse(method_name == "gsgcs", "#FCE3E3", "#E3F0FA")
  p <- p + annotate("rect", xmin = 0, xmax = 10, ymin = n_rows + 1.35, ymax = n_rows + 2.20, fill = header_fill, colour = NA)
  # Method title
  p <- p + annotate("text", x = 5, y = n_rows + 1.77, label = paste0(method_display, " | ", cancer_name), size = 4.3, fontface = "bold")
  # Column header
  header_y <- n_rows + 0.88
  p <- p + annotate("text", x = 1.55, y = header_y, label = "Functional cluster", size = 3, fontface = "bold") + 
    annotate("text", x = 4.15, y = header_y, label = "OS", size = 3, fontface = "bold") + 
    annotate("text", x = 5.05, y = header_y, label = "PFI", size = 3, fontface = "bold") + 
    annotate("text", x = 5.95, y = header_y, label = "PCAWG", size = 3, fontface = "bold") + 
    annotate("text", x = 7.45, y = header_y, label = "Shared all 3", size = 2.7, fontface = "bold", lineheight = 0.9) + 
    annotate("text", x = 9.1, y = header_y, label = "Total", size = 3, fontface = "bold")
  # Rows
  for (i in seq_len(n_rows)) {
    row_data <- top[i, , drop = FALSE]
    y <- row_data$row
    # alternating background
    if (i %% 2 == 0) {
      p <- p + annotate("rect", xmin = 0, xmax = 10, ymin = y - 0.42, ymax = y + 0.42, fill = "grey97", colour = NA)
    }
    # Functional cluster
    cluster_label <- as.character(row_data$Functional_cluster)
    # shorten very long names
    cluster_label <- str_wrap(cluster_label, width = 25)
    p <- p + annotate("text", x = 0.12, y = y, label = cluster_label, hjust = 0, size = 2.65, lineheight = 0.9)
    # Endpoint values
    endpoint_values <- c(as.numeric(row_data$os), as.numeric(row_data$pfi), as.numeric(row_data$time))
    endpoint_x <- c(4.15, 5.05, 5.95)
    for (j in seq_along(endpoint_values)) {
      value <- endpoint_values[j]
      if (!is.na(value) && value > 0) {
        point_fill <- ifelse(method_name == "gsgcs", "#F8766D", "#619CFF")
        p <- p + annotate("point", x = endpoint_x[j], y = y, shape = 21, size = 3.2, stroke = 0.65, fill = point_fill, colour = "grey20")
      } else {
        p <- p + annotate("point", x = endpoint_x[j], y = y, shape = 21, size = 2.7, stroke = 0.5, fill = "white", colour = "grey70")
      }
    }
    # Shared all 3
    p <- p + annotate("text", x = 7.45, y = y, label = as.character(row_data$shared_all3), size = 3)
    # Total
    p <- p + annotate("text", x = 9.1, y = y, label = as.character(row_data$total_unique), size = 3)
  }
  # Horizontal lines
  for (y_line in seq(0.5, n_rows + 0.5, by = 1)) {
    p <- p + annotate("segment", x = 0, xend = 10, y = y_line, yend = y_line, colour = "grey85", linewidth = 0.35)
  }
  p + theme_void() + theme(plot.margin = margin(2, 0, 2, 0))
}
# 21. Legend panel
make_legend_panel <- function(count_breaks) {
  p <- ggplot() + xlim(0, 15) + ylim(0, 5)
  # Functional category
  p <- p + annotate("text", x = 0.2, y = 4.55, label = "Node color: KEGG functional category (Level 1)", hjust = 0, size = 3.3, fontface = "bold")
  legend_categories <- names(category_colors)
  if (length(legend_categories) > 0) {
    category_y <- seq(3.9, 0.7, length.out = length(legend_categories))
    for (i in seq_along(legend_categories)) {
      category <- legend_categories[i]
      p <- p + annotate("point", x = 0.45, y = category_y[i], shape = 21, size = 3.2, fill = category_colors[category], colour = "grey20", stroke = 0.6) + annotate("text", x = 0.8, y = category_y[i], label = category, hjust = 0, size = 2.8)
    }
  }
  # Node size
  p <- p + annotate("text", x = 7.0, y = 4.55, label = "Node size:\nPrioritized genes in pathway", hjust = 0, vjust = 1, size = 3.2, fontface = "bold", lineheight = 0.95)
  size_values <- count_breaks
  if (length(size_values) > 4) {
    index <- unique(round(seq(1, length(size_values), length.out = 4)))
    size_values <- size_values[index]
  }
  if (length(size_values) > 0) {
    size_y <- seq(3.3, 0.9, length.out = length(size_values))
    size_min <- min(size_values)
    size_max <- max(size_values)
    for (i in seq_along(size_values)) {
      value <- size_values[i]
      if (size_max == size_min) {
        point_size <- 4
      } else {
        point_size <- 3 + 5 * (value - size_min) / (size_max - size_min)
      }
      p <- p + annotate("point", x = 7.5, y = size_y[i], shape = 21, size = point_size, fill = "grey90", colour = "grey25", stroke = 0.7) + annotate("text", x = 8.15, y = size_y[i], label = as.character(value), hjust = 0, size = 3)
    }
  }
  # Recurrence
  p <- p + annotate("text", x = 10.1, y = 4.55, label = "Node border: endpoint recurrence", hjust = 0, size = 3.3, fontface = "bold")
  recurrence_labels <- c("Present in 1 endpoint\n(endpoint-specific)", "Present in 2 endpoints", "Present in all 3 endpoints\n(OS + PFI + PCAWG)")
  recurrence_y <- c(3.45, 2.25, 1.05)
  recurrence_stroke <- c(0.3, 0.9, 1.6)
  recurrence_colour <- c("grey70", "grey35", "black")
  for (i in seq_along(recurrence_labels)) {
    p <- p + annotate("point", x = 10.55, y = recurrence_y[i], shape = 21, size = 6, fill = "grey90", colour = recurrence_colour[i], stroke = recurrence_stroke[i]) + annotate("text", x = 11.1, y = recurrence_y[i], label = recurrence_labels[i], hjust = 0, size = 2.8, lineheight = 0.95)
  }
  # Separators
  p <- p + annotate("segment", x = 6.6, xend = 6.6, y = 0.3, yend = 4.7, colour = "grey85", linewidth = 0.5) + annotate("segment", x = 9.7, xend = 9.7, y = 0.3, yend = 4.7, colour = "grey85", linewidth = 0.5)
  p + theme_void() + theme(panel.border = element_rect(colour = "grey85", fill = NA, linewidth = 0.6), plot.margin = margin(5, 5, 5, 5))
}

# 22. Generate one cancer figure
make_cancer_functional_figure <- function(cancer_name) {
  cat("\n============================================\n")
  cat("Generating: ", cancer_name, "\n")
  cat("============================================\n")
  net_coxcs <- build_union_network(cancer_name, "coxcs")
  net_gsgcs <- build_union_network(cancer_name, "gsgcs")
  all_counts <- c(net_coxcs$nodes$Count, net_gsgcs$nodes$Count)
  all_counts <- all_counts[is.finite(all_counts)]
  if (length(all_counts) == 0) {
    count_limits <- c(1, 10)
  } else {
    count_limits <- range(all_counts, na.rm = TRUE)
    if (count_limits[1] == count_limits[2]) {
      count_limits <- c(count_limits[1], count_limits[1] + 1)
    }
  }
  count_breaks <- pretty(count_limits, n = 4)
  count_breaks <- count_breaks[count_breaks >= count_limits[1] & count_breaks <= count_limits[2]]
  if (length(count_breaks) == 0) {
    count_breaks <- mean(count_limits)
  }
  p_coxcs <- plot_union_network(net_coxcs, cancer_name, "coxcs", count_limits, count_breaks)
  p_gsgcs <- plot_union_network(net_gsgcs, cancer_name, "gsgcs", count_limits, count_breaks)
  p_legend <- make_legend_panel(count_breaks)
  cluster_coxcs <- make_cluster_summary(cancer_name, "coxcs")
  cluster_gsgcs <- make_cluster_summary(cancer_name, "gsgcs")
  table_coxcs <- plot_cluster_table(cluster_coxcs, cancer_name, "coxcs")
  table_gsgcs <- plot_cluster_table(cluster_gsgcs, cancer_name, "gsgcs")
  network_row <- p_coxcs | p_gsgcs
  table_row <- table_coxcs | table_gsgcs
  final_plot <- (network_row / p_legend / table_row) +plot_layout(heights = c(5.2, 2.0, 3.0)) +plot_annotation(title = "KEGG pathway recurrence across OS, PFI and PCAWG",theme = theme(plot.title = element_text(size = 19,face = "bold",hjust = 0.5,margin = margin(b = 10))))
  png_file <- file.path(output_dir, paste0("Functional_recurrence_", cancer_name, ".png"))
  ggsave(filename = png_file, plot = final_plot, width = figure_width, height = figure_height, units = "in", dpi = 600, bg = "white", limitsize = FALSE)
  pdf_file <- file.path(output_dir, paste0("Functional_recurrence_", cancer_name, ".pdf"))
  ggsave(filename = pdf_file, plot = final_plot, width = figure_width, height = figure_height, units = "in", device = cairo_pdf, bg = "white", limitsize = FALSE)
  write.csv(cluster_coxcs, file.path(output_dir, paste0(cancer_name, "_CoxCS_functional_clusters.csv")), row.names = FALSE)
  write.csv(cluster_gsgcs, file.path(output_dir, paste0(cancer_name, "_GSGCS_functional_clusters.csv")), row.names = FALSE)
  summary_out <- bind_rows(get_method_summary(cancer_name, "coxcs"), get_method_summary(cancer_name, "gsgcs"))
  write.csv(summary_out, file.path(output_dir, paste0(cancer_name, "_recurrence_summary.csv")), row.names = FALSE)
  cat("Saved:\n", png_file, "\n", pdf_file, "\n")
  return(list(figure = final_plot, CoxCS_network = net_coxcs, GSGCS_network = net_gsgcs, CoxCS_clusters = cluster_coxcs, GSGCS_clusters = cluster_gsgcs, summary = summary_out))
}


# 23. First test OV only
# Strongly recommended:
# run OV first before all five cancers

test_OV <-make_cancer_functional_figure("OV")

# 24. Check OV summary
print(test_OV$summary,n = Inf, width = Inf)
print(test_OV$CoxCS_clusters,n = Inf,width = Inf)
print(test_OV$GSGCS_clusters,n = Inf, width = Inf)

# 25. If OV works:
# generate all five cancers
all_cancer_figures <- setNames(map(cancer_list,make_cancer_functional_figure), cancer_list)

# 26. Combined summary
all_summary <- map_dfr(cancer_list, function(cancer_name) {bind_rows(get_method_summary(cancer_name,"coxcs"),get_method_summary(cancer_name,"gsgcs"))})
print(all_summary,n=Inf,width=Inf)
write.csv(all_summary,file.path(output_dir,"ALL_cancers_recurrence_summary.csv"),row.names=FALSE)

# 27. Check output files
output_check <- map_dfr(cancer_list, function(cancer_name) {
  png_file <- file.path(output_dir, paste0("Functional_recurrence_", cancer_name, ".png"))
  pdf_file <- file.path(output_dir, paste0("Functional_recurrence_", cancer_name, ".pdf"))
  tibble(cancer = cancer_name, PNG = file.exists(png_file), PDF = file.exists(pdf_file), PNG_MB = ifelse(file.exists(png_file), round(file.info(png_file)$size / 1024^2, 2), NA_real_))
})
print(output_check, n = Inf)




#### （3.6）富集结果绘制 ： figure 5 ----
# Hallmark / KEGG
library(readxl)
library(dplyr)
library(stringr)
library(igraph)
library(ggraph)
library(ggplot2)
library(ggforce)
library(patchwork)
library(tidyr)
library(grid)
library(cowplot)

# USER SETTINGS
# DATABASE <- "hallmark"
DATABASE <- "kegg"

TOP_N <- 80
PADJ <- 0.05
JACCARD_CUTOFF <- 0.05

# Load enrichment results
pathway <- read_excel("富集通路-new1.xlsx")

# Data preparation
df <- pathway %>%
  filter( by == DATABASE,Padj < PADJ,Count >=3) %>%
  mutate(type=tolower(type),geneID=as.character(geneID)) %>%
  group_by(type) %>%
  arrange(desc(LOGP)) %>%
  slice_head(n=TOP_N) %>%
  ungroup()

# Functions
split_gene <- function(x){unique(unlist(str_split(x,"/")))}

jaccard <- function(a,b){length(intersect(a,b)) /length(union(a,b))}

# Build network
build_network <- function(data,method){
  dat <- data %>%filter(type==method) %>%distinct(ID,.keep_all=TRUE)
  gene_list <- lapply(dat$geneID,split_gene)
  names(gene_list) <- dat$ID
  edges <- data.frame()
  for(i in 1:(length(gene_list)-1)){
    for(j in (i+1):length(gene_list)){
      sim <- jaccard(gene_list[[i]],gene_list[[j]])
      if(sim >= JACCARD_CUTOFF){
        edges <- rbind(edges,data.frame(from=names(gene_list)[i],
                                        to=names(gene_list)[j],
                                        weight=sim))}}}
  nodes <- dat %>% 
    select(ID,Description,LOGP,Level_1_Category) %>%
    mutate(
      Description = gsub("^HALLMARK_", "", Description),
      Description = gsub("_", " ", Description)
    ) %>%
    rename(name=ID)
  g <- graph_from_data_frame(edges,vertices=nodes,directed=FALSE)
  return(g)
  
}


g_gsgcs <- build_network(df,"gsgcs")
g_coxcs <- build_network(df,"coxcs")

# Community detection
add_module <- function(g){
  clu <- cluster_louvain(g, weights=E(g)$weight)
  V(g)$module <- as.factor(membership(clu))
  return(g)
}

g_gsgcs <- add_module(g_gsgcs)
g_coxcs <- add_module(g_coxcs)

library(RColorBrewer)
module_num_gsgcs <- length(unique(V(g_gsgcs)$module))
module_num_coxcs <- length(unique(V(g_coxcs)$module))

# palette_gsgcs <- colorRampPalette(RColorBrewer::brewer.pal(12,"Set3"))(module_num_gsgcs)
# palette_coxcs <- colorRampPalette(RColorBrewer::brewer.pal(12,"Pastel1"))(module_num_coxcs)
palette_gsgcs <- colorRampPalette(
  c("#1B9E77","#ABDDA4","#66C2A5","#3288BD","#91BFDB",
    "#8C510A","#B35806","#FDAE61",    
    "#A50026","#C51B7D","#E78AC3","#5E4FA2"))(module_num_gsgcs)

palette_coxcs <- colorRampPalette(
  c(
    "#D73027","#F46D43","#FC8D59","#FEE090",
    "#313695","#762A83",
    "#2C7BB6","#4575B4","#74ADD1","#006837"))(module_num_gsgcs)


# Plot network

# hallmark
# > size_range <- range(c(V(g_gsgcs)$LOGP,V(g_coxcs)$LOGP))
# > size_range
# [1]  2.185862 20.754276
# > size_range <- range(c(V(g_gsgcs)$LOGP))
# > size_range
# [1]  3.226359 20.754276
# > size_range <- range(c(V(g_coxcs)$LOGP))
# > size_range
# [1]  2.185862 14.063146

# > edge_range <- range(c(E(g_gsgcs)$weight,E(g_coxcs)$weight))
# > edge_range
# [1] 0.0500000 0.3506494
# > edge_range <- range(c(E(g_gsgcs)$weight))
# > edge_range
# [1] 0.0500000 0.1717172
# > edge_range <- range(c(E(g_coxcs)$weight))
# > edge_range
# [1] 0.05050505 0.35064935

# kegg
# > size_range <- range(c(V(g_gsgcs)$LOGP,V(g_coxcs)$LOGP))
# > size_range
# [1]  12.81869 125.26627
# > size_range <- range(c(V(g_gsgcs)$LOGP))
# > size_range
# [1]  32.83737 125.26627
# > size_range <- range(c(V(g_coxcs)$LOGP))
# > size_range
# [1]  12.81869 105.76465

# > edge_range <- range(c(E(g_gsgcs)$weight,E(g_coxcs)$weight))
# > edge_range
# [1] 0.050000 0.705036
# > edge_range <- range(c(E(g_gsgcs)$weight))
# > edge_range
# [1] 0.05345912 0.70503597
# > edge_range <- range(c(E(g_coxcs)$weight))
# > edge_range
# [1] 0.0500000 0.5287356
GLOBAL_SIZE_RANGE <- range(
  c(
    V(g_gsgcs)$LOGP,
    V(g_coxcs)$LOGP
  )
)
plot_network <- function(g,title,module_palette,database,GLOBAL_SIZE_RANGE){
  
  if(database=="hallmark"){
    
    size_limits <- GLOBAL_SIZE_RANGE # c(0,20)
    size_range <- c(3,12)
    size_breaks <- c(5,10,15,20)
    
    edge_limits <- c(0.05,0.35)
    edge_breaks <- c(0.05,0.10,0.20,0.30)
    
  }else if(database=="kegg"){
    
    size_limits <- GLOBAL_SIZE_RANGE # range(V(g)$LOGP)
    size_range <- c(3,12)
    size_breaks <- c(20,50,100)
    
    edge_limits <- c(0.05,0.70)
    edge_breaks <- c(0.05,0.20,0.40,0.60)
    
  }
  # calculate node connectivity
  # g <- delete_vertices(g,V(g)[degree(g)==0])
  V(g)$degree <- degree(g)
  label_threshold <- 1
  V(g)$label_show <- ifelse(V(g)$degree > label_threshold,V(g)$name,"")
  # V(g)$label_show <- ifelse(V(g)$degree >= 2 | V(g)$LOGP >= quantile(V(g)$LOGP,0.7),V(g)$Description,"")
  ggraph(g,layout="fr")+
    geom_mark_hull(aes(x=x,y=y,group=module,fill=module),
                   colour=NA,alpha=0.08,expand=unit(8,"mm"))+
    geom_edge_link(aes(width=weight),
                   colour="grey60",alpha=0.25)+
    geom_node_point(aes(size=LOGP,fill=module),
                    colour="black",shape=21,stroke=0.3,max.overlaps=100)+
    geom_node_text(aes(label=Description),# label_show),# 
                   repel=TRUE,size=3)+
    scale_edge_width(limits=edge_limits,
                     range=c(0.3,2),
                     breaks=edge_breaks,
                     name="Gene overlap")+
    scale_size_continuous(limits=size_limits,
                          range=size_range,
                          breaks=size_breaks,
                          name="-log10(Padj)")+
    scale_fill_manual(values=module_palette,guide="none")+
    theme_void()+
    ggtitle(title) +
    theme(plot.title= element_text(size=16,face="bold",hjust=0.5))
}

pA <- plot_network(g_gsgcs,"GSGCS",palette_gsgcs,DATABASE,GLOBAL_SIZE_RANGE )
pB <- plot_network(g_coxcs,"CoxCS",palette_coxcs,DATABASE,GLOBAL_SIZE_RANGE )
# legend <- cowplot::get_legend(pA)
# pA_no_legend <- pA + theme(legend.position="none")
# Network statistics
# 修改核心部分
network_stats <- function(g,data,method){
  ## pathway数量
  pathway_num <-data %>%filter( type==method) %>% distinct(ID) %>%nrow()
  ## 所有可能pathway pair
  total_pairs <-pathway_num *(pathway_num-1) /2
  ## 实际存在shared gene连接pair
  shared_pairs <-ecount(g)
  Shared_pathway_pairs <-shared_pairs /total_pairs *100
  ## LCC
  comp <- components(g)
  LCC <-max(comp$csize) /vcount(g) *100
  ## density
  Density <-edge_density(g)
  ## modularity
  Modularity <-modularity(cluster_louvain(g,weights=E(g)$weight))
  Mean_degree <- mean(degree(g))
  Weighted_degree <- mean(strength(g,weights=E(g)$weight))
  # Pathway redundancy index
  n_pathway <- igraph::vcount(g)
  n_edge <- igraph::ecount(g)
  redundancy_index <- ifelse(n_pathway > 1,n_edge / (n_pathway * (n_pathway - 1) / 2),0)
  data.frame(Method=ifelse(method=="gsgcs","GSGCS","CoxCS"),
             Number_of_pathways=pathway_num,
             Shared_pathway_pairs=Shared_pathway_pairs,
             LCC=LCC,
             Mean_degree=Mean_degree,
             Weighted_degree=Weighted_degree,
             Redundancy_index = redundancy_index,
             # Density=Density,
             Modularity=Modularity)
}

stats <- bind_rows(network_stats(g_gsgcs,df,"gsgcs"),
                   network_stats(g_coxcs,df,"coxcs" ))

print(stats)

# Panel C
plot_stats <-stats %>%pivot_longer(cols=-Method,
                                   names_to="Metric",
                                   values_to="Value")
# 根据数据库选择展示指标
if(DATABASE=="hallmark"){
  
  metric_order <- c("Shared_pathway_pairs",
                    "LCC",
                    "Redundancy_index",
                    "Modularity")
  metric_labels <- c("Shared pathway pairs (%)",
                     "Largest connected component (%)",
                     "Pathway overlap index",
                     "Modularity")
}else if(DATABASE=="kegg"){
  metric_order <- c("Number_of_pathways",
                    "Shared_pathway_pairs",
                    "LCC",
                    "Redundancy_index",
                    "Modularity")
  metric_labels <- c("Number of pathways",
                     "Shared pathway pairs (%)",
                     "Largest connected component (%)",
                     "Pathway overlap index",
                     "Modularity")
}
plot_stats <- plot_stats %>% filter(Metric %in% metric_order)
plot_stats$Metric <- factor(plot_stats$Metric,
                            levels=metric_order,labels=metric_labels)


pC <-ggplot(plot_stats,aes(Method,Value,fill=Method))+
  geom_col(width=0.65)+
  geom_text(aes(label=sprintf("%.2f",Value)),vjust=-0.3,size=4)+
  facet_wrap(~Metric,scales="free_y",nrow=1)+
  scale_fill_manual(values=c(CoxCS="#9DC3E6",GSGCS="#EFA3A3"))+
  theme_classic()+
  theme(legend.position="none",strip.text=element_text(face="bold"))

top_panel <- (pA | pB | guide_area()) + 
  plot_layout(guides="collect", widths=c(1,1,0.35)) & theme(legend.position="right")
Figure5 <- top_panel / pC + plot_layout( heights=c(1.25,1))

# Save
ggsave(paste0("Figure5_",DATABASE,"_final.png"),
       Figure5,width=16,height=10,dpi=300,bg="white")



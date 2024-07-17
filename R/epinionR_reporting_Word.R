#' Generate Word report
#'
#' This function allows users to generate a simple Epinion Standard Word report based on
#' a prepared Word template and an OrderForm.
#'
#' @param df datainput (dataframe).
#' @param orderform_folder directory where orderforms stored. Default is the current working dir.
#' @param output_folder directory where output report should be placed. Default is the current working dir.
#' @param listOrderForm list of Orderform.
#' @param ... Additional arguments may apply.
#' @return Epinion Standard Word report
#' @examples
#' # Set the example folder of the package as the working directory
#' setwd(system.file("extdata", package = "epinionPS"))
#'
#' # Read the data
#' mydata <- epinion_read_data(file = example("sample_test_Word_report.sav"))
#'
#' # State the default template to generate report
#' report_template <- "Tabular_Landscape_Template_20221116.docx"
#'
#' # Get all Word OrderForms in the folder
#' list_orderForm <- list.files(path = paste0(getwd(), "/"), pattern='OrderForm_Word_*')
#'
#' # Run reports
#' epinion_reporting_Word(mydata,
#'                        orderform_folder = paste0(getwd(), "/"),
#'                        output_folder = paste0(getwd(), "/"),
#'                        list_orderForm)

epinion_reporting_Word = function(df,
                                  orderform_folder = getwd(),
                                  output_folder = getwd(),
                                  listOrderForm) {
  if (!require(pacman)) install.packages("pacman")
  pacman::p_load(dplyr,
                tools,
                devtools,
                anesrake,
                haven,
                openxlsx,
                expss,
                labelled,
                lubridate,
                janitor,
                officer)

  if (any(grepl("~$OrderForm_Word", listOrderForm, fixed = TRUE)) == TRUE) {
    stop("The orderform is opening by Excel application. Please close the file before running the syntax.")
  }

  for (orderForm in listOrderForm) {
    # df <- haven::as_factor(df, only_labelled = TRUE)

    if (!file.exists(paste0(orderform_folder, orderForm))) {
      stop(paste0("'", paste0(orderform_folder, orderForm), "'", " does not exist in the working folder."))

    } else {
      start_time <- Sys.time()

      epinion_create_Word_report_main(df,
                                      orderform_folder = orderform_folder,
                                      output_folder = output_folder,
                                      orderForm)

      end_time <- Sys.time()
      print(end_time - start_time)

    }

  }

}

epinion_create_Word_report_main = function(df,
                                           orderform_folder,
                                           output_folder,
                                           report_orderform){

  # 01. Read information in the order form
  ## Sheet "GeneralInformation"
  report_generalInfor <- openxlsx::readWorkbook(report_orderform,
                                                sheet = "GeneralInformation",
                                                colNames = TRUE,
                                                rowNames = FALSE,
                                                skipEmptyRows = FALSE,
                                                check.names = FALSE,
                                                na.strings = ""
                                                )

  report_output <- paste0(report_generalInfor[1, 2], ".docx")
  client_name <- report_generalInfor[2, 2]
  report_title <- report_generalInfor[4, 2]
  report_date <- format(Sys.Date(), "%d.%b %Y")
  project_number <- as.character(report_generalInfor[1, 3])

  no_interview <- as.character(dim(df)[1])
  project_type <- report_generalInfor[3, 2]

  if (file.exists(report_output) ) {
    out_temp = try(file(report_output, open = "w"),
                   silent = TRUE)

    if (suppressWarnings("try-error" %in% class(out_temp))) {
      stop(paste0("The output Word report '", report_output,  "' is being opened by Word application. Please close the file before running the script."))

    }

    close(out_temp)

  }

  ## Sheet "Analysis"
  report_analysis <- openxlsx::readWorkbook(report_orderform,
                                            sheet = "Analysis",
                                            colNames = TRUE,
                                            rowNames = FALSE,
                                            skipEmptyRows = FALSE,
                                            check.names = FALSE,
                                            na.strings = ""
                                            )

  ## Check if variables in orderform exist in the dataframe before generating the report
  var_list = ""
  for (i in report_analysis$Independent.variables){
    var_list = ifelse(is.na(i), var_list, paste(var_list, i, sep=","))

  }

  for (i in report_analysis$Dependent.variables){
    var_list = ifelse(is.na(i), var_list, paste(var_list, i, sep=","))

  }

  var_list = ifelse(is.na(report_generalInfor[6, 2]), var_list, paste(var_list, report_generalInfor[6, 2], sep=","))

  var_list = substr(var_list,2, nchar(var_list))
  var_list = strsplit(var_list, "[\\+,]+")[[1]]
  varNotInData = var_list[unlist(!(var_list %in% names(df)))]

  if (length(varNotInData) > 0) {
    stop(paste0("Variable '", unlist(varNotInData), "' is not in the dataset. Please review the orderform.", "\n  "))

  }


  var_list[unlist(!(var_list %in% names(df)))]

  if (length(which(!is.na(report_analysis$Sig_test))) > 0) {
    sigtestTest <- "I krydstabuleringerne er der vist et signifikansniveau for beregningen. Beregningen er foretaget på baggrund af en Pearson Chi-Square. Når signifikansniveauet er under 0,05 er beregningen signifikant. Epinion"

  } else {
    sigtestTest <- "Epinion"

  }

  report_analysis$Section <- factor(report_analysis$Section, levels=unique(report_analysis$Section))

  section <- report_analysis %>%
    group_by(Section) %>%
    summarise(count=n()) %>%
    mutate(start = NA,
           end = NA)

  for (i in 1:dim(section)[1]) {
    if (i==1) {
      section$start[i] = 1
      section$end[i] = section$start[i] + section$count[i]-1

    } else {
      section$start[i] = section$end[i-1] + 1
      section$end[i] = section$start[i] + section$count[i]-1

    }

  }

  # 02. Create PPT
  options(dplyr.summarise.inform = FALSE)

  for (i in 1:dim(section)[1]) {

    ### Add section/chapter
    if (i==1) {
      report_output <- epinion_addsection(report_template,
                                          section_name = section$Section[i],
                                          target = report_output)

    } else {
      report_output <- epinion_addsection(report_output,
                                           section_name = section$Section[i],
                                           target = report_output)

    }

    # ### Add tables in each section

    docx_obj <- read_docx(report_output)

    for (j in section$start[i]:section$end[i]) {
      ## Get the common information
      # weight var
      if (report_generalInfor[5, 2] == "No") {
        weight_var = "totalt"

      } else if (report_generalInfor[5, 2] == "Yes") {
        weight_var = report_generalInfor[6, 2]

      }

      # weighted base: to determine if we would like to report weighted or unweighted base
      if (is.na(report_generalInfor[7, 2]) == FALSE & report_generalInfor[7, 2] == "No") {
        base = "base_chart_unweighted"

      } else {
        base = "base_chart"

      }

      # filter
      if (is.na(report_analysis$`Filter/Condition.(If.requried)`[j])){
        filter = "totalt=1"

      } else {
        filter = report_analysis$`Filter/Condition.(If.requried)`[j]

      }

      # counted_value (Multiple variables)
      if (is.na(report_analysis$Counted_value[j])) {
        counted_value = "Yes" #1

      } else {
        counted_value = report_analysis$Counted_value[j]

      }

      # sort
      if (!is.na(report_analysis$Sort_by[j])) {
        sort_by_cat <- report_analysis$Sort_by[j]

      } else {
        sort_by_cat <- 'default'

      }

      if (!is.na(report_analysis$Sort_order[j]) & report_analysis$Sort_order[j] == "Ascending") {
        sort_order <- FALSE

      } else if (!is.na(report_analysis$Sort_order[j]) & report_analysis$Sort_order[j] == "Descending") {
        sort_order <- TRUE

      } else {
        sort_order <- report_analysis$Sort_order[j]

      }

      # add_base
      add_base <- report_analysis$Add_base[j]

      # Sigtest
      if (is.na(report_analysis$Sig_test[j])) {
        sigtest = FALSE

      } else {
        sigtest = TRUE

      }

      # significance level
      if (!is.na(report_analysis$Sig_level[j])){
        sig_level <- as.numeric(report_analysis$Sig_level[j])

      } else {
        sig_level <- 0.05

      }

      # Adjust bonferroni
      if (!is.na(report_analysis$Adjust_bonferroni[j]) & report_analysis$Adjust_bonferroni[j] == "Yes"){
        bonferroni <- TRUE

      } else {
        bonferroni <- FALSE

      }

      # Table title
      tbl_title = report_analysis$Table.Title[j]
      rp_analysis = report_analysis$Analysis[j]

      ## Frequency of single variable
      if (rp_analysis == "Freq" & is.na(report_analysis$Advanced.Option[j])){
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_addtbl_freq_single(df, col_var, row_var, weight_var,
                                     tbl_title = tbl_title,
                                     filter = filter,
                                     add_base, base,
                                     docx_obj = docx_obj,
                                     rp_analysis = rp_analysis
                                     )
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Frequency of multiple variables
      if (rp_analysis == "Multiple Freq (By Case)"){
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_addtbl_freq_multiple(df, col_var, row_var, weight_var,
                                       tbl_title = tbl_title,
                                       counted_value = counted_value,
                                       filter = filter,
                                       add_base, base,
                                       sort_by_cat = sort_by_cat, sort_order = sort_order,
                                       docx_obj = docx_obj,
                                       rp_analysis = rp_analysis
                                       )
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Frequency of grid variables
      if (rp_analysis == "Freq" & is.na(report_analysis$Advanced.Option[j]) == FALSE){
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_addtbl_freq_grid(df, col_var, row_var, weight_var,
                                   tbl_title = tbl_title,
                                   filter = filter,
                                   add_base, base,
                                   sort_by_cat = sort_by_cat, sort_order = sort_order,
                                   docx_obj = docx_obj,
                                   rp_analysis = rp_analysis
                                   )
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Frequency of grid variables with mean
      if (rp_analysis == "Freq Grid with Mean" & is.na(report_analysis$Advanced.Option[j]) == FALSE){
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_addtbl_freq_grid(df, col_var, row_var, weight_var,
                                   tbl_title = tbl_title,
                                   filter = filter,
                                   add_base, base,
                                   sort_by_cat = sort_by_cat, sort_order = sort_order,
                                   docx_obj = docx_obj,
                                   rp_analysis = rp_analysis
                                   )
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Crosstab: single vs. single
      if (rp_analysis == "Crosstabs"){
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_addtbl_crosstab_single(df, col_var, row_var, weight_var,
                                         tbl_title = tbl_title,
                                         filter = filter,
                                         add_base, base,
                                         sigtest = sigtest,
                                         sig_level = sig_level, bonferroni = bonferroni,
                                         docx_obj = docx_obj,
                                         rp_analysis = rp_analysis
                                         )
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Crosstab: multiple
      if (rp_analysis == "Multiple Crosstabs (By Case)"){
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_addtbl_crosstab_multiple(df, col_var, row_var, weight_var,
                                           tbl_title = tbl_title,
                                           counted_value = counted_value,
                                           filter = filter,
                                           add_base, base,
                                           sort_by_cat = sort_by_cat, sort_order = sort_order,
                                           sigtest = sigtest,
                                           sig_level = sig_level, bonferroni = bonferroni,
                                           docx_obj = docx_obj,
                                           rp_analysis = rp_analysis
                                           )
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Crosstab grid with mean
      if (rp_analysis == "Crosstabs Row with Mean" |
          rp_analysis == "Multiple Crosstabs Row (By Case) with Mean") {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_addtbl_crosstab_grid_with_mean(df, col_var, row_var, weight_var,
                                                 tbl_title = tbl_title,
                                                 filter = filter,
                                                 add_base, base,
                                                 counted_value = counted_value,
                                                 sigtest = sigtest,
                                                 sig_level = sig_level, bonferroni = bonferroni,
                                                 docx_obj = docx_obj,
                                                 rp_analysis = rp_analysis
                                                 )
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Open answer
      if (rp_analysis == "Open answer"){
        report_mode = "unique"

        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_add_openans(df, row_var,
                              filter = filter,
                              report_mode = report_mode,
                              tbl_title = tbl_title,
                              docx_obj = docx_obj,
                              rp_analysis = rp_analysis
                              )
        )

        print_status(rp_analysis, row_var, col_var)

      } else if (rp_analysis == "Open answer - Full") {
        report_mode = "full"

        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        docx_obj <- catchErrorAndRetry(
          epinion_add_openans(df, row_var,
                              filter = filter,
                              report_mode = report_mode,
                              tbl_title = tbl_title,
                              docx_obj = docx_obj,
                              rp_analysis = rp_analysis
                              )
        )

        print_status(rp_analysis, row_var, col_var)

      }


    }
    print(paste0("================ Please wait for section '", section$Section[i] ,"' to be saved ================"))
    print(docx_obj, target = report_output)
  }

  print("================ Please wait for the report to be saved ʢ´• ᴥ •`ʡ ================")

  # Update slide table of content and cover slide

  docx_obj <- read_docx(report_output)
  docx_obj <- docx_obj %>%
    cursor_reach(keyword = "REPORT TITLE") %>%
    body_replace_all_text(old_value = "REPORT TITLE", new_value = report_title,
                          only_at_cursor = TRUE) %>%
    cursor_reach(keyword = "CLIENT NAME") %>%
    body_replace_all_text(old_value = "CLIENT NAME", new_value = client_name,
                          only_at_cursor = TRUE) %>%
    cursor_reach(keyword = "REPORT DATE") %>%
    body_replace_all_text(old_value = "REPORT DATE", new_value = report_date,
                          only_at_cursor = TRUE) %>%
    cursor_reach(keyword = "PROJECT NUMBER") %>%
    body_replace_all_text(old_value = "PROJECT NUMBER", new_value = project_number,
                          only_at_cursor = TRUE) %>%
    cursor_reach(keyword = "<number of interviews>") %>%
    body_replace_all_text(old_value = "<number of interviews>", new_value = no_interview,
                          only_at_cursor = TRUE) %>%
    cursor_reach(keyword = "<webcatitext>") %>%
    body_replace_all_text(old_value = "<webcatitext>", new_value = project_type,
                          only_at_cursor = FALSE, fixed = TRUE) %>%
    cursor_begin()%>%
    cursor_reach(keyword = "text_significanttest") %>%
    # docx_show_chunk()%>%
    body_replace_all_text(old_value = "text_significanttest", new_value = sigtestTest,
                          only_at_cursor = FALSE, fixed = FALSE) %>%
    cursor_reach(keyword = "<client name>") %>%
    body_replace_all_text(old_value = "<client name>", new_value = client_name,
                          only_at_cursor = FALSE) %>%
    # docx_show_chunk() %>%
    cursor_reach(keyword = "Epinion står naturligvis") %>%
    # docx_show_chunk()%>%
    body_replace_all_text(old_value = "Epinion står naturligvis", new_value = " står naturligvis",
                          only_at_cursor = TRUE, fixed = FALSE) %>%
    print(report_output)

}


# ##############################################################################
# Add table to the report
# ##############################################################################
# Prepare frequency table
epinion_calc_freq = function(x, col_var, row_var, weight_var,
                            filter_var, filter_val,
                            rp_analysis,
                            counted_value = "Yes") {

  df_tbl <- x %>%
    mutate(totalt = 1) %>%
    mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
    filter(filter_tmp == filter_val)

  if (rp_analysis == "Freq" & length(row_var) == 1) {
    df_tbl <- df_tbl %>%
      filter(!is.na(!!as.symbol(row_var))) %>%
      tab_cells("|" = unvr(..$row_var ))  %>%
      tab_cols(total(label = "Total")) %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_cpct(total_statistic = "w_cpct", total_label = "Total") %>%
      tab_stat_cases(total_label = "Total", total_statistic = c("w_cases"), label = "Count") %>%
      tab_weight(weight = totalt) %>%
      tab_stat_cases(total_label = "Total", total_statistic = c("u_cases"), label = "Count_unweight") %>%
      tab_pivot(stat_position = "outside_columns") %>%
      if_na(0) %>%
      as.data.frame(.) %>%
      mutate(Pct_str = paste0(round_half_up(Total,0), "%"),
             Procentandel = paste0(Pct_str, "\\n", paste0("(", round_half_up(`Total|Count`,0), ")")),
             Procentandel_unweighted = paste0(Pct_str, "\\n", paste0("(", round_half_up(`Total|Count_unweight`,0), ")")),
             row_labels = ifelse(row_labels == "#Total", "Total", row_labels))

  } else if ((rp_analysis == "Freq" & length(row_var) > 1) |
             rp_analysis == "Freq Grid with Mean") {
    df_tbl <- df_tbl %>%
      haven::as_factor(., only_labelled = TRUE) %>%
      mutate(weight_variable = !!as.symbol(weight_var)) %>%
      select(., all_of(row_var), weight_variable) %>%
      stack_with_labels() %>%
      tab_cells(value) %>%
      tab_cols(variable) %>%
      tab_weight(weight = weight_value) %>%
      tab_stat_cpct(total_statistic = "w_cpct", total_label = "Total") %>%
      tab_stat_cases(total_label = "Total", total_statistic = c("w_cases"), label = "Count") %>%
      tab_weight(weight = total()) %>%
      tab_stat_cases(total_label = "Total", total_statistic = c("u_cases"), label = "Count_unweight") %>%
      tab_pivot() %>%
      if_na(0) %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      separate(row_labels, c("row_varOrder", "row_labels"), sep = "\\|", remove = TRUE) %>%
      mutate(row_varOrder = as.numeric(row_varOrder)) %>%
      arrange(row_varOrder)

  } else if (rp_analysis == "Multiple Freq (By Case)") {
    names_old <- unique(unlist(as_factor(x)[, row_var]))
    names_new <- ifelse(names_old == counted_value, 1, 0)
    df_conversion <- data.frame(names_old, names_new)

    df_tbl <- df_tbl %>%
      as_factor(., only_labelled = TRUE) %>%
      mutate(across(!!as.symbol(head(row_var, 1)) : !!as.symbol(tail(row_var, 1)),
                    ~ deframe(df_conversion)[.])) %>%
      set_variable_labels(., .labels = var_label(x)) %>%
      filter(if_any(head(row_var, 1):tail(row_var, 1), ~. == 1)) %>%
      tab_cells(mdset(..$head(row_var, 1) %to% ..$tail(row_var, 1)))  %>%
      tab_cols(total(label = "Total")) %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_cpct(total_statistic = "w_cpct", total_label = "Total") %>%
      tab_stat_cases(total_label = "Total", total_statistic = c("w_cases"), label = "Count") %>%
      tab_weight(weight = totalt) %>%
      tab_stat_cases(total_label = "Total", total_statistic = c("u_cases"), label = "Count_unweight") %>%
      tab_pivot(stat_position = "outside_columns") %>%
      if_na(0) %>%
      as.data.frame(.) %>%
      mutate(Total = ifelse(row_labels == "#Total", sum(Total, na.rm = TRUE)-100, Total)) %>%
      mutate(Pct_str = paste0(round_half_up(Total,0), "%"),
             Procentandel = paste0(Pct_str, "\\n", paste0("(", round_half_up(`Total|Count`,0), ")")),
             Procentandel_unweighted = paste0(Pct_str, "\\n", paste0("(", round_half_up(`Total|Count_unweight`,0), ")")),
             row_labels = ifelse(row_labels == "#Total", "Total", row_labels))


  }

  df_tbl

}


# Prepare mean table
epinion_calc_mean = function(x, col_var, row_var, weight_var,
                            filter_var, filter_val,
                            rp_analysis) {


  if (grepl("with Mean", rp_analysis, fixed = TRUE) > 0) {
    df_tbl <- x %>%
      mutate(totalt = 1) %>%
      mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
      filter(filter_tmp == filter_val) %>%
      haven::as_factor(., only_labelled = TRUE) %>%
      mutate(weight_variable = !!as.symbol(weight_var)) %>%
      select(., all_of(row_var), weight_variable) %>%
      stack_with_labels() %>%
      tab_cells(value) %>%
      tab_cols(variable) %>%
      tab_weight(weight = weight_value) %>%
      tab_stat_mean(label = "Gns.") %>%
      tab_pivot() %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      separate(row_labels, c("row_varOrder", "row_labels"), sep = "\\|", remove = TRUE) %>%
      mutate(`Gns.` = round_half_up(`Gns.`, 2)) %>%
      select(., -row_varOrder)

  }

  df_tbl

}

# Prepare crosstab table
epinion_calc_crosstab = function(x, col_var, row_var, weight_var,
                                 filter_var, filter_val,
                                 rp_analysis,
                                 counted_value = "Yes",
                                 sigtest = FALSE,
                                 sig_level = 0.05, bonferroni = FALSE) {

  df_tbl_input <- x %>%
    mutate(totalt = 1) %>%
    mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
    filter(filter_tmp == filter_val)

  if (length(col_var) == 1 & length(row_var) == 1) {
    df_tbl_input <- df_tbl_input %>%
      filter(!is.na(!!as.symbol(col_var))) %>%
      select(., totalt, !!as.symbol(weight_var), !!as.symbol(row_var), !!as.symbol(col_var)) %>%
      tab_cells("|" = unvr(..$row_var ))  %>%
      tab_cols("|" = unvr(..$col_var ), total(label = "Total"))

  } else if (length(col_var) > 1 & length(row_var) == 1) {
    names_old <- unique(unlist(as_factor(x)[, col_var]))
    names_new <- ifelse(names_old == counted_value, 1, 0)
    df_conversion <- data.frame(names_old, names_new)

    var_multiple_col_tmp <- df_tbl_input %>%
      as_factor(., only_labelled = TRUE) %>%
      mutate(id_tmp = row.names(.)) %>%
      select(., id_tmp, all_of(col_var)) %>%
      mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

    df_tbl_input <- df_tbl_input %>%
      select(., !all_of(col_var)) %>%
      mutate(id_tmp = row.names(.)) %>%
      left_join(., var_multiple_col_tmp, by = "id_tmp") %>%
      select(., -id_tmp) %>%
      set_variable_labels(., .labels = var_label(x)) %>%
      select(., totalt, !!as.symbol(weight_var), !!as.symbol(row_var), all_of(col_var)) %>%
      filter(if_any(head(col_var, 1):tail(col_var, 1), ~. == 1)) %>%
      tab_cells("|" = unvr(..$row_var ))  %>%
      tab_cols(mdset(..$head(col_var, 1) %to% ..$tail(col_var, 1)), total(label = "Total"))

  } else if (length(col_var) == 1 & length(row_var) > 1) {
    names_old <- unique(unlist(as_factor(x)[, row_var]))
    names_new <- ifelse(names_old == counted_value, 1, 0)
    df_conversion <- data.frame(names_old, names_new)

    var_multiple_row_tmp <- df_tbl_input %>%
      as_factor(., only_labelled = TRUE) %>%
      mutate(id_tmp = row.names(.)) %>%
      select(., id_tmp, all_of(row_var)) %>%
      mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

    df_tbl_input <- df_tbl_input %>%
      select(., !all_of(row_var)) %>%
      mutate(id_tmp = row.names(.)) %>%
      left_join(., var_multiple_row_tmp, by = "id_tmp") %>%
      select(., -id_tmp) %>%
      set_variable_labels(., .labels = var_label(x)) %>%
      select(., totalt, !!as.symbol(weight_var), !!as.symbol(col_var), all_of(row_var)) %>%
      filter(if_any(head(row_var, 1):tail(row_var, 1), ~. == 1)) %>%
      filter(!is.na(!!as.symbol(col_var))) %>%
      tab_cells(mdset(..$head(row_var, 1) %to% ..$tail(row_var, 1))) %>%
      tab_cols("|" = unvr(..$col_var ), total(label = "Total"))

  } else if (length(col_var) > 1 & length(row_var) > 1) {
    names_old <- unique(unlist(as_factor(x)[, row_var]))
    names_new <- ifelse(names_old == counted_value, 1, 0)
    df_conversion <- data.frame(names_old, names_new)

    var_multiple_row_tmp <- df_tbl_input %>%
      as_factor(., only_labelled = TRUE) %>%
      mutate(id_tmp = row.names(.)) %>%
      select(., id_tmp, all_of(row_var)) %>%
      mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

    var_multiple_col_tmp <- df_tbl_input %>%
      as_factor(., only_labelled = TRUE) %>%
      mutate(id_tmp = row.names(.)) %>%
      select(., id_tmp, all_of(col_var)) %>%
      mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

    df_tbl_input <- df_tbl_input %>%
      select(., !all_of(col_var)) %>%
      select(., !all_of(row_var)) %>%
      mutate(id_tmp = row.names(.)) %>%
      left_join(., var_multiple_col_tmp, by = "id_tmp") %>%
      left_join(., var_multiple_row_tmp, by = "id_tmp") %>%
      select(., -id_tmp) %>%
      set_variable_labels(., .labels = var_label(x)) %>%
      select(., totalt, !!as.symbol(weight_var), all_of(row_var), all_of(col_var)) %>%
      filter(if_any(head(col_var, 1):tail(col_var, 1), ~. == 1)) %>%
      tab_cells(mdset(..$head(row_var, 1) %to% ..$tail(row_var, 1))) %>%
      tab_cols(mdset(..$head(col_var, 1) %to% ..$tail(col_var, 1)), total(label = "Total"))

  }

  df_tbl <- df_tbl_input %>%
    tab_weight(weight = ..$weight_var) %>%
    tab_stat_cpct(total_statistic = "w_cpct", total_label = "Total") %>%
    tab_stat_cases(total_label = "Total", total_statistic = c("w_cases"), label = "Count") %>%
    tab_weight(weight = totalt) %>%
    tab_stat_cases(total_label = "Total", total_statistic = c("u_cases"), label = "Count_unweight") %>%
    tab_pivot(stat_position = "outside_columns") %>%
    if_na(0) %>%
    as.data.frame(.)

  if(!sigtest) {
    out <- df_tbl

  } else {
    if (length(col_var) == 1) {
      out <- df_tbl_input %>%
        tab_cols("|" = unvr(..$col_var ))

    } else if (length(col_var) > 1) {
      out <- df_tbl_input %>%
        tab_cols(mdset(..$head(col_var, 1) %to% ..$tail(col_var, 1)))

    }

    out <- out %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_cpct(total_label = "Total") %>%
      # tab_stat_cpct(total_label = "Total", total_statistic = "w_cases") %>%
      tab_last_sig_cpct(sig_level = sig_level, bonferroni = bonferroni, digits = 0,
                        subtable_marks = "greater", sig_labels = LETTERS) %>%
      tab_last_round(digits = 0) %>%
      tab_pivot() %>%
      if_na(0) %>%
      as.data.frame(.) %>%
      left_join(., select(df_tbl, row_labels, Total), by = "row_labels") %>%
      mutate(Total = round_half_up(Total, 0)) %>%
      epinion_add_percent()

    # replace blank with 0%
    out[out == ""] <- "0%"

    names(out) <- ifelse(!names(out) %in% c("row_labels", "Total"),
                        paste0(gsub("\\|", "\\\\n (", names(out)), ")"),
                        names(out))

  }

  out

}

# Prepare crosstab mean table
epinion_calc_mean_crosstab = function(x, col_var, row_var, weight_var,
                                     filter_var, filter_val,
                                     rp_analysis,
                                     counted_value = "Yes",
                                     sigtest = FALSE,
                                     sig_level = 0.05, bonferroni = FALSE) {
  df_tbl_input <- x %>%
    mutate(totalt = 1) %>%
    mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
    filter(filter_tmp == filter_val)

  if (length(col_var) == 1) {
    df_tbl_input <- df_tbl_input %>%
      filter(!is.na(!!as.symbol(col_var))) %>%
      tab_cells("|" = unvr(..$row_var ))  %>%
      tab_cols("|" = unvr(..$col_var ), total(label = "Total"))

  } else if (length(col_var) > 1) {
    names_old <- unique(unlist(as_factor(x)[, col_var]))
    names_new <- ifelse(names_old == counted_value, 1, 0)
    df_conversion <- data.frame(names_old, names_new)

    var_multiple_col_tmp <- df_tbl_input %>%
      as_factor(., only_labelled = TRUE) %>%
      mutate(id_tmp = row.names(.)) %>%
      select(., id_tmp, all_of(col_var)) %>%
      mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

    df_tbl_input <- df_tbl_input %>%
      select(., !all_of(col_var)) %>%
      mutate(id_tmp = row.names(.)) %>%
      left_join(., var_multiple_col_tmp, by = "id_tmp") %>%
      select(., -id_tmp) %>%
      set_variable_labels(., .labels = var_label(x)) %>%
      select(., totalt, !!as.symbol(weight_var), !!as.symbol(row_var), all_of(col_var)) %>%
      filter(if_any(head(col_var, 1):tail(col_var, 1), ~. == 1)) %>%
      tab_cells("|" = unvr(..$row_var ))  %>%
      tab_cols(mdset(..$head(col_var, 1) %to% ..$tail(col_var, 1)), total(label = "Total"))

  }

  df_tbl <- df_tbl_input %>%
    tab_weight(weight = ..$weight_var) %>%
    tab_stat_mean(label = "Gns.") %>%
    tab_pivot() %>%
    as.data.frame(.) %>%
    select(., !row_labels) %>%
    pivot_longer(., everything(), names_to = "row_labels", values_to = "Gns.") %>%
    mutate(`Gns.` = round_half_up(`Gns.`, 2)) %>%
    tab_transpose() %>%
    as.data.frame(.) %>%
    mutate(row_labels = row.names(.))

  names(df_tbl) <- df_tbl[1,]

  df_tbl <- df_tbl %>%
    select(., row_labels, colnames(.)) %>%
    filter(row_labels != "row_labels")

  if (!sigtest) {
    out <- df_tbl

  } else {
    if (length(col_var) == 1) {
      out <- df_tbl_input %>%
        tab_cols("|" = unvr(..$col_var))

    } else if (length(col_var) > 1) {
      out <- df_tbl_input %>%
        tab_cols(mdset(..$head(col_var, 1) %to% ..$tail(col_var, 1)))

    }

    out <- out %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_mean_sd_n() %>%
      tab_last_sig_means(sig_level = sig_level, bonferroni = bonferroni,
                         digits = 2, subtable_marks = "greater") %>%
      tab_pivot() %>%
      as.data.frame(.) %>%
      filter(row_labels == "Mean") %>%
      select(., -row_labels) %>%
      pivot_longer(., everything(), names_to = "row_labels", values_to = "Gns.") %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      mutate(row_labels = row.names(.))

    names(out) <- out[1,]

    out <- out %>%
      select(., row_labels, colnames(.)) %>%
      filter(row_labels != "row_labels") %>%
      left_join(., select(df_tbl, row_labels, Total), by = "row_labels")

    names(out) <- ifelse(!names(out) %in% c("row_labels", "Total"),
                         paste0(gsub("\\|", "\\\\n (", names(out)), ")"),
                         names(out))

  }

  out

}

# ==============================================================================
# Add frequency table of single variable to Word report
epinion_addtbl_freq_single = function(x, col_var, row_var, weight_var,
                                      tbl_title = empty_content(),
                                      filter = "totalt=1",
                                      add_base, base,
                                      docx_obj, rp_analysis) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_tbl_new <- epinion_calc_freq(x, col_var, row_var, weight_var,
                                  filter_var, filter_val,
                                  rp_analysis)

  if (is.na(add_base)) {
    df_tbl <- df_tbl_new %>%
      dplyr::select(., row_labels, Pct_str) %>%
      rename(" " = row_labels,
             Procentandel = Pct_str)

  } else if (is.na(add_base) == FALSE & base == "base_chart_unweighted") {
    df_tbl <- df_tbl_new %>%
      dplyr::select(., row_labels, Procentandel_unweighted) %>%
      rename(" " = row_labels,
             Procentandel = Procentandel_unweighted)

  } else if (is.na(add_base) == FALSE & base == "base_chart") {
    df_tbl <- df_tbl_new %>%
      dplyr::select(., row_labels, Procentandel) %>%
      rename(" " = row_labels)

  }

  docx_obj <- epinion_addtbl(docx_obj, df_tbl, tbl_title)

  docx_obj

}

# ==============================================================================
# Add frequency table of multiple variable to Word report
epinion_addtbl_freq_multiple = function(x, col_var, row_var, weight_var,
                                        tbl_title = empty_content(),
                                        counted_value = "Yes",
                                        filter = "totalt=1",
                                        add_base, base,
                                        sort_by_cat = 'default', sort_order = TRUE,
                                        docx_obj, rp_analysis) {

  var_multiple <- strsplit(row_var, "\\+")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_tbl_new <- epinion_calc_freq(x, col_var, var_multiple, weight_var,
                                 filter_var, filter_val,
                                 rp_analysis,
                                 counted_value)

  # base
  if (is.na(add_base)) {
    df_tbl <- df_tbl_new %>%
      dplyr::select(., row_labels, Pct_str, Total) %>%
      rename(Procentandel = Pct_str)

  } else if (is.na(add_base) == FALSE & base == "base_chart_unweighted") {
    df_tbl <- df_tbl_new %>%
      dplyr::select(., row_labels, Procentandel_unweighted, Total) %>%
      rename(Procentandel = Procentandel_unweighted)

  } else if (is.na(add_base) == FALSE & base == "base_chart") {
    df_tbl <- df_tbl_new %>%
      dplyr::select(., row_labels, Procentandel, Total)

  }

  # sort option
  if (sort_by_cat == 'default') {
    df_tbl <- df_tbl %>%
      select(., -Total) %>%
      rename(" " = row_labels)

  } else {
    excluded_cats <- append(strsplit(sort_by_cat, "\\,")[[1]], "Total")

    df_sort_0 <- filter(df_tbl, row_labels %in% excluded_cats)
    df_sort_1 <- filter(df_tbl, !row_labels %in% excluded_cats)
    df_sort_1 <- df_sort_1[order(df_sort_1$Total, decreasing = sort_order), ]
    df_tbl <- bind_rows(df_sort_1, df_sort_0) %>%
      select(., -Total) %>%
      rename(" " = row_labels)

  }

  docx_obj <- epinion_addtbl(docx_obj, df_tbl, tbl_title)

  docx_obj

}


# ==============================================================================
# Add frequency table of grid variable to Word report
epinion_addtbl_freq_grid = function(x, col_var, row_var, weight_var,
                                    tbl_title = empty_content(),
                                    filter = "totalt=1",
                                    add_base, base,
                                    sort_by_cat = 'default', sort_order = TRUE,
                                    docx_obj, rp_analysis) {

  var_grid <- strsplit(row_var, "\\,")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_tbl_new <- epinion_calc_freq(x, col_var, var_grid, weight_var,
                                 filter_var, filter_val,
                                 rp_analysis)

  # sort
  if (sort_by_cat == 'default') {
    df_tbl_new <- df_tbl_new

  } else {
    if (sort_order == TRUE) {
      df_tbl_new <- df_tbl_new %>%
        arrange(desc(!!as.symbol(sort_by_cat)))

    } else {
      df_tbl_new <- df_tbl_new %>%
        arrange(!!as.symbol(sort_by_cat))

    }

  }

  df_tbl_count <- df_tbl_new %>%
    select(., row_labels, ends_with("Count")) %>%
    rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
    mutate(across(where(is.numeric), round_half_up))

  df_tbl_count_unweight <- df_tbl_new %>%
    select(., row_labels, ends_with("Count_unweight")) %>%
    rename_with(~ sub("\\|Count_unweight$", "", .x), everything())

  out <- df_tbl_new %>%
    select(., row_labels, !contains("Count"), - row_varOrder) %>%
    mutate(across(where(is.numeric), round_half_up)) %>%
    epinion_add_percent()

  out_weight <- out
  out_weight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                        out_weight[-1], df_tbl_count[-1])

  out_unweight <- out
  out_unweight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                          out_unweight[-1], df_tbl_count_unweight[-1])


  # base
  if (is.na(add_base)) {
    df_tbl <- out %>%
      rename("Total" = "#Total")

  } else if (is.na(add_base) == FALSE & base == "base_chart_unweighted") {
    df_tbl <- out_unweight %>%
      rename("Total" = "#Total")

  } else if (is.na(add_base) == FALSE & base == "base_chart") {
    df_tbl <- out_weight %>%
      rename("Total" = "#Total")

  }

  if (rp_analysis == "Freq Grid with Mean") {
    var_grid_mean <- paste0(var_grid, "_mean")

    df_tbl_new_mean <- epinion_calc_mean(x, col_var, var_grid_mean, weight_var,
                                        filter_var, filter_val,
                                        rp_analysis)

    df_tbl <- df_tbl %>%
      left_join(., df_tbl_new_mean, by = "row_labels") %>%
      rename(" " = row_labels)

  } else {
    df_tbl <- df_tbl %>%
      rename(" " = row_labels)

  }

  docx_obj <- epinion_addtbl(docx_obj, df_tbl, tbl_title)

  docx_obj

}

# ==============================================================================
# Add crosstab table of single variable to Word report
epinion_addtbl_crosstab_single = function(x, col_var, row_var, weight_var,
                                          tbl_title = empty_content(),
                                          filter = "totalt=1",
                                          add_base, base,
                                          sigtest = FALSE,
                                          sig_level = 0.05, bonferroni = FALSE,
                                          docx_obj, rp_analysis) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_tbl_new <- epinion_calc_crosstab(x, col_var, row_var, weight_var,
                                     filter_var, filter_val,
                                     rp_analysis)

  df_tbl_count <- df_tbl_new %>%
    select(., row_labels, ends_with("Count")) %>%
    rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
    mutate(across(where(is.numeric), round_half_up))

  df_tbl_count_unweight <- df_tbl_new %>%
    select(., row_labels, ends_with("Count_unweight")) %>%
    rename_with(~ sub("\\|Count_unweight$", "", .x), everything())

  out <- df_tbl_new %>%
    select(., row_labels, !contains("Count")) %>%
    mutate(across(where(is.numeric), round_half_up)) %>%
    epinion_add_percent()

  out_weight <- out
  out_weight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                        out_weight[-1], df_tbl_count[-1])

  out_unweight <- out
  out_unweight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                          out_unweight[-1], df_tbl_count_unweight[-1])

  if (sigtest) {
    df_tbl_new_sigtest <- epinion_calc_crosstab(x, col_var, row_var, weight_var,
                                                filter_var, filter_val,
                                                rp_analysis,
                                                sigtest = sigtest,
                                                sig_level = sig_level,
                                                bonferroni = bonferroni)

    total_row <- out %>%
      filter(row_labels == "#Total")

    names(total_row) <- names(df_tbl_new_sigtest)

    df_tbl_new_sigtest <- df_tbl_new_sigtest %>%
      filter(row_labels != "#Total") %>%
      bind_rows(., total_row)

    out_weight <- df_tbl_new_sigtest
    out_weight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                          out_weight[-1], df_tbl_count[-1])

    out_unweight <- df_tbl_new_sigtest
    out_unweight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                            out_unweight[-1], df_tbl_count_unweight[-1])

  }

  # base
  if (is.na(add_base)) {
    df_tbl <- out %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      rename(" " = row_labels)

  } else if (is.na(add_base) == FALSE & base == "base_chart_unweighted") {
    df_tbl <- out_unweight %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      rename(" " = row_labels)

  } else if (is.na(add_base) == FALSE & base == "base_chart") {
    df_tbl <- out_weight %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      rename(" " = row_labels)

  }

  docx_obj <- epinion_addtbl(docx_obj, df_tbl, tbl_title)

  docx_obj

}

# ==============================================================================
# Add crosstab table of single variable to Word report
epinion_addtbl_crosstab_grid_with_mean = function(x, col_var, row_var, weight_var,
                                                        tbl_title = empty_content(),
                                                        filter = "totalt=1",
                                                        add_base, base,
                                                        counted_value = "Yes",
                                                        sigtest = FALSE,
                                                        sig_level = 0.05, bonferroni = FALSE,
                                                        docx_obj, rp_analysis) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  var_grid_mean <- paste0(row_var, "_mean")
  var_multiple_col <- strsplit(col_var, "\\+")[[1]]

  df_tbl_new <- epinion_calc_crosstab(x, var_multiple_col, row_var, weight_var,
                                      filter_var, filter_val,
                                      rp_analysis,
                                      counted_value)

  df_tbl_mean <- epinion_calc_mean_crosstab(x, var_multiple_col, var_grid_mean, weight_var,
                                            filter_var, filter_val,
                                            rp_analysis,
                                            counted_value)



  df_tbl_count <- df_tbl_new %>%
    select(., row_labels, ends_with("Count")) %>%
    rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
    mutate(across(where(is.numeric), round_half_up))

  names(df_tbl_count) <- NULL

  df_tbl_count_unweight <- df_tbl_new %>%
    select(., row_labels, ends_with("Count_unweight")) %>%
    rename_with(~ sub("\\|Count_unweight$", "", .x), everything())

  names(df_tbl_count_unweight) <- NULL

  out <- df_tbl_new %>%
    select(., row_labels, !contains("Count")) %>%
    mutate(across(where(is.numeric), round_half_up)) %>%
    epinion_add_percent()

  out_weight <- out
  out_weight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                        out_weight[-1], df_tbl_count[-1])

  out_unweight <- out
  out_unweight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                          out_unweight[-1], df_tbl_count_unweight[-1])

  if (sigtest) {
    df_tbl_new_sigtest <- epinion_calc_crosstab(x, var_multiple_col, row_var, weight_var,
                                                filter_var, filter_val,
                                                rp_analysis,
                                                counted_value,
                                                sigtest = sigtest,
                                                sig_level = sig_level,
                                                bonferroni = bonferroni)

    df_tbl_mean_sigtest <- epinion_calc_mean_crosstab(x, var_multiple_col, var_grid_mean, weight_var,
                                                      filter_var, filter_val,
                                                      rp_analysis,
                                                      counted_value,
                                                      sigtest = sigtest,
                                                      sig_level = sig_level,
                                                      bonferroni = bonferroni)

    total_row <- out %>%
      filter(row_labels == "#Total")

    names(total_row) <- names(df_tbl_new_sigtest)

    df_tbl_new_sigtest <- df_tbl_new_sigtest %>%
      filter(row_labels != "#Total") %>%
      bind_rows(., total_row)

    out_weight <- df_tbl_new_sigtest
    out_weight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                          out_weight[-1], df_tbl_count[-1])

    out_unweight <- df_tbl_new_sigtest
    out_unweight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                            out_unweight[-1], df_tbl_count_unweight[-1])

  }

  # base
  if (is.na(add_base)) {
    df_tbl <- out

  } else if (is.na(add_base) == FALSE & base == "base_chart_unweighted") {
    df_tbl <- out_unweight

  } else if (is.na(add_base) == FALSE & base == "base_chart") {
    df_tbl <- out_weight

  }

  if (!sigtest) {
    df_tbl <- df_tbl %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      bind_rows(., df_tbl_mean) %>%
      rename(" " = row_labels)

  } else {
    df_tbl <- df_tbl %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      bind_rows(., df_tbl_mean_sigtest) %>%
      rename(" " = row_labels)

  }

  docx_obj <- epinion_addtbl(docx_obj, df_tbl, tbl_title)

  docx_obj

}

# ==============================================================================
# Add crosstab table of multiple variable to Word report
epinion_addtbl_crosstab_multiple = function(x, col_var, row_var, weight_var,
                                            tbl_title = empty_content(),
                                            counted_value = "Yes",
                                            filter = "totalt=1",
                                            add_base, base,
                                            sort_by_cat = 'default', sort_order = TRUE,
                                            sigtest = FALSE,
                                            sig_level = 0.05, bonferroni = FALSE,
                                            docx_obj, rp_analysis) {

  var_multiple_row <- strsplit(row_var, "\\+")[[1]]
  var_multiple_col <- strsplit(col_var, "\\+")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_tbl_new <- epinion_calc_crosstab(x, var_multiple_col, var_multiple_row, weight_var,
                                     filter_var, filter_val,
                                     rp_analysis,
                                     counted_value = counted_value)

  if (length(var_multiple_col) == 1 & length(var_multiple_row) > 1) {
    # sort
    if (sort_by_cat == 'default') {
      df_tbl_new <- df_tbl_new

    } else {
      excluded_cats <- append(strsplit(sort_by_cat, "\\,")[[1]], "#Total")

      df_sort_0 <- filter(df_tbl_new, row_labels %in% excluded_cats)
      df_sort_1 <- filter(df_tbl_new, !row_labels %in% excluded_cats)
      df_sort_1 <- df_sort_1[order(df_sort_1$Total, decreasing = sort_order), ]
      df_tbl_new <- bind_rows(df_sort_1, df_sort_0)

    }

    out <- df_tbl_new %>%
      select(., row_labels, !contains("Count")) %>%
      filter(., row_labels != "#Total") %>%
      bind_rows(summarise(., across(where(is.numeric), sum),
                          across(where(is.character), ~'#Total'))) %>%
      mutate(across(where(is.numeric), round_half_up)) %>%
      epinion_add_percent()

  } else if (length(var_multiple_col) > 1 & length(var_multiple_row) == 1) {
    out <- df_tbl_new %>%
      select(., row_labels, !contains("Count")) %>%
      mutate(across(where(is.numeric), round_half_up)) %>%
      epinion_add_percent()


  } else if (length(var_multiple_col) > 1 & length(var_multiple_row) > 1) {
    out <- df_tbl_new %>%
      select(., row_labels, !contains("Count")) %>%
      filter(., row_labels != "#Total") %>%
      bind_rows(summarise(., across(where(is.numeric), sum),
                          across(where(is.character), ~'#Total'))) %>%
      mutate(across(where(is.numeric), round_half_up)) %>%
      epinion_add_percent()

  }

  df_tbl_count <- df_tbl_new %>%
    select(., row_labels, ends_with("Count")) %>%
    rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
    mutate(across(where(is.numeric), round_half_up))

  df_tbl_count_unweight <- df_tbl_new %>%
    select(., row_labels, ends_with("Count_unweight")) %>%
    rename_with(~ sub("\\|Count_unweight$", "", .x), everything())

  out_weight <- out
  out_weight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                        out_weight[-1], df_tbl_count[-1])

  out_unweight <- out
  out_unweight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                          out_unweight[-1], df_tbl_count_unweight[-1])

  if (sigtest) {
    df_tbl_new_sigtest <- epinion_calc_crosstab(x, var_multiple_col, var_multiple_row, weight_var,
                                                filter_var, filter_val,
                                                rp_analysis,
                                                counted_value,
                                                sigtest = sigtest,
                                                sig_level = sig_level,
                                                bonferroni = bonferroni)

    total_row <- out %>%
      filter(row_labels == "#Total")

    names(total_row) <- names(df_tbl_new_sigtest)

    df_tbl_new_sigtest <- df_tbl_new_sigtest %>%
      filter(row_labels != "#Total") %>%
      bind_rows(., total_row)

    # Sort
    df_tbl_new_sigtest <- df_tbl_new_sigtest[match(unique(df_tbl_new$row_labels), df_tbl_new_sigtest$row_labels),]

    out_weight <- df_tbl_new_sigtest
    out_weight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                          out_weight[-1], df_tbl_count[-1])

    out_unweight <- df_tbl_new_sigtest
    out_unweight[-1] <- Map(function(x, y) sprintf("%0s\\n(%d)", x, y),
                            out_unweight[-1], df_tbl_count_unweight[-1])

  }

  # base
  if (is.na(add_base)) {
    df_tbl <- out %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      rename(" " = row_labels)

  } else if (is.na(add_base) == FALSE & base == "base_chart_unweighted") {
    df_tbl <- out_unweight %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      rename(" " = row_labels)

  } else if (is.na(add_base) == FALSE & base == "base_chart") {
    df_tbl <- out_weight %>%
      mutate(row_labels = ifelse(row_labels == "#Total", "Total", row_labels)) %>%
      rename(" " = row_labels)

  }

  docx_obj <- epinion_addtbl(docx_obj, df_tbl, tbl_title)

  docx_obj

}


# ##############################################################################
# Add open answers to the report
# ##############################################################################
epinion_add_openans = function(x, open_ans_var,
                               filter = "totalt=1", report_mode,
                               tbl_title,
                               docx_obj, rp_analysis) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  if (report_mode == "unique") {
    open_answer <- x %>%
      mutate(totalt = 1) %>%
      mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
      filter(filter_tmp == filter_val) %>%
      arrange(!!as.symbol(open_ans_var)) %>%
      dplyr::select(., !!as.symbol(open_ans_var)) %>%
      filter(!!as.symbol(open_ans_var) != "") %>%
      group_by(!!as.symbol(open_ans_var)) %>%
      summarise()

  } else if (report_mode == "full") {
    open_answer <- x %>%
      mutate(totalt = 1) %>%
      mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
      filter(filter_tmp == filter_val) %>%
      arrange(!!as.symbol(open_ans_var)) %>%
      dplyr::select(., !!as.symbol(open_ans_var)) %>%
      filter(!!as.symbol(open_ans_var) != "")

  }

  landscape_one_column <- block_section(
    prop_section(page_size = page_size(orient = "landscape"),
                 type = "continuous"
                 )
    )

  landscape_two_columns <- block_section(
    prop_section(page_size = page_size(orient = "landscape"),
                 type = "continuous",
                 # page_margins = page_mar(top = 2),
                 section_columns = section_columns(widths = c(5.22, 5.22))
                 )
    )

  docx_obj <- docx_obj %>%
    # body_add_break() %>%
    epinion_body_add_OA_legend(legend = tbl_title,
                               legend_name = "Åbne Besvarelser",
                               legend_style = getOption("crosstable_style_legend", "Style3"),
                               seqfield = "",
                               par_after = FALSE)

  docx_obj <- body_end_block_section(docx_obj, value = landscape_one_column)
  # there stops section with landscape_one_column

  # there starts section with landscape_two_columns
  for (text in open_answer[[open_ans_var]]) {
    docx_obj <- docx_obj %>%
      body_add_par(value = text, style = "Style4", pos = "after")

  }
  docx_obj <- body_add_break(docx_obj)

  docx_obj <- body_end_block_section(docx_obj, value = landscape_two_columns)
  # there stops section with landscape_two_columns

  docx_obj

}


# ##############################################################################
# Function to add table to the report
# ##############################################################################
epinion_addsection = function(docx_obj, section_name = empty_content(), target){

  docx_obj <- read_docx(docx_obj) %>%
    body_end_section_continuous()
    # body_add_break()

  # add section
  section <- block_caption(section_name, style = "heading 1")
  docx_obj <- body_add_caption(docx_obj, section, pos = "on")

  print(docx_obj, target = target)

}

pacman::p_load(crosstable)

epinion_addtbl = function(docx_obj, tbl_data, tbl_title = empty_content()){

  docx_obj <- docx_obj %>%
    body_add_table_legend(legend = tbl_title,
                          legend_name = "Tabel",
                          name_format = fp_text(color = "#0f283c", font.size = 12, bold=TRUE),
                          legend_style = getOption("crosstable_style_legend", "Style3")
    ) %>%
    body_add_table(tbl_data, header = TRUE, style = "Standardtabel",
                   pos = "after", align_table = "left",
                   alignment = c("l", rep("c", length(tbl_data)-1))
    ) %>%
    body_add_break()%>%
    body_end_section_continuous() #%>%
    # body_add_break()

  docx_obj

}


# ##############################################################################
# Other functions
# ##############################################################################
# Print status of report processing in the console
print_status = function(rp_analysis, row_var, col_var){
  print(paste0(rp_analysis, " - ",
               row_var, " - ",
               col_var, " - Done!"))
}

# Add percent sign (%) to expss ctables
epinion_add_percent = function(x, digits = get_expss_digits(), excluded_rows = "count", ...){
  nas = is.na(x)
  x[nas] = ""

  cols_idx = 2:dim(x)[2]

  for (col in cols_idx) {
    for (row in 1:dim(x)[1]){
      if (!grepl(excluded_rows, x[row, 1], perl = TRUE)){
        if (suppressWarnings(is.na(as.numeric(as.character(x[row,col]))))) {
          x[row,col] = sub(" ", "% ", trimws(x[row,col]))

        } else {
          x[row,col] = paste0(trimws(x[row,col]), "%")

        }

      }

    }

  }

  x <- x[!grepl("Std. dev.", x$row_labels),]
  x <- x[!grepl("Unw. valid N", x$row_labels),]
  x

}


# Add legend for Open answers
epinion_body_add_legend = function (doc, legend, legend_name, bookmark, legend_prefix,
                                    legend_style, name_format, seqfield, style, legacy) {
  if (packageVersion("officer") < "0.4" || legacy) {
    if (!legacy) {
      warn("You might want to update officer to v0.4+ in order to get the best of crosstable::body_add_xxx_legend().",
           .frequency = "once", .frequency_id = "body_add_xxx_legend_officer_version")

    }

    if (is_missing(style)) {
      style = getOption("crosstable_style_strong", "strong")

    }

    rtn = body_add_legend_legacy(doc = doc, legend = legend,
                                 legend_name = legend_name, bookmark = bookmark, legend_style = legend_style,
                                 style = style, seqfield = seqfield)

    return(rtn)

  }

  if (lifecycle::is_present(style)) {
    lifecycle::deprecate_warn("0.2.2", "body_add_X_legend(style)", "body_add_X_legend(name_format)",
                   details = "Therefore, its value has been ignored. Use `legacy=TRUE` to override.")

  }

  legend = paste0(legend_prefix, legend)
  fp_text2 = officer::fp_text_lite

  if (is.null(name_format)) {
    name_format = getOption("crosstable_format_legend_name",
                            fp_text2(bold = TRUE))

  }

  fp_size = fp_text2(font.size = name_format$font.size)
  legend = glue::glue(legend, .envir = parent.frame())
  # legend_name = paste0(legend_name, " ")
  bkm = run_word_field(seqfield, prop = name_format)
  if (!is.null(bookmark)) {
    bkm = run_bookmark(bookmark, bkm)

  }

  legend_fpar = fpar(ftext(legend_name, name_format), bkm,
                     ftext(": ", name_format), ftext(legend, fp_size))
  body_add_fpar(doc, legend_fpar, style = legend_style)

}

epinion_body_add_OA_legend = function (doc, legend, ..., bookmark = NULL, legend_style = getOption("crosstable_style_legend",
                                       doc$default_styles$paragraph), style = deprecated(), legend_prefix = NULL,
                                      name_format = NULL, legend_name = "Figure", seqfield = "SEQ Figure \\* Arabic",
                                      par_after = FALSE, legacy = FALSE) {
  # crosstable:::check_dots_empty()
  if (missing(par_after))
    par_after = getOption("crosstable_figure_legend_par_after",
                          FALSE)

  if (missing(legend_prefix))
    legend_prefix = getOption("crosstable_figure_legend_prefix",
                              NULL)

  doc = epinion_body_add_legend(doc = doc, legend = legend, legend_name = legend_name,
                        bookmark = bookmark, legend_prefix = legend_prefix, legend_style = legend_style,
                        name_format = name_format, seqfield = seqfield,
                        legacy = legacy)

  if (par_after) {
    doc = body_add_normal(doc, "")

  }

  doc
}

# Stack table
stack_with_labels = function(df, cols = NULL) {
  if(is.function(cols)) {
    cols = cols(colnames(df))

  }
  if(is.null(cols)) cols = TRUE
  need_cols = df[,cols]
  weight_cols = df[, "weight_variable"]
  all_var_labs = lapply(need_cols[,-which(names(need_cols) %in% c("weight_variable"))], var_lab)
  # check for empty labels
  no_var_lab = lengths(all_var_labs) == 0
  all_var_labs[no_var_lab] = "|"
  all_var_labs = rep(unlist(all_var_labs), each = nrow(df))
  # need_cols = lapply(need_cols[,-length(need_cols)], as.labelled) %>%
  #   as.data.frame(.)
  need_cols = need_cols[,-length(need_cols)] %>%
    as.data.frame(.)

  value = do.call(c, need_cols)
  weight_value = rep(do.call(c, weight_cols), length(need_cols))
  res = data.frame(variable = all_var_labs, value = value, weight_value = weight_value) %>%
    mutate(var_name = row.names(.)) %>%
    mutate(var_order = sapply(map(str_extract_all(var_name, "[0-9]+"), as.numeric), sum)) %>%
    group_by(variable) %>%
    mutate(var_order = min(var_order)) %>%
    ungroup() %>%
    mutate(variable = paste0(var_order, "|", variable))

  var_lab(res) = "|"
  res

}

# Catch error and retry
# To cope with the situation when the random error 'set_val_lab' occurs
# https://github.com/gdemin/expss/issues/107
# The author does not yet find out the root cause as well as the solution for it

catchErrorAndRetry <- function(fnc) {
  c = 0 # set counter to zero

  repeat {
    error <- FALSE

    result <- suppressWarnings(
      tryCatch(fnc,
               error = function(e) { error <<- TRUE})
    )

    if(error == FALSE) { break }

    if(c == 10) { break}

    c = c + 1

    print(paste0("***** Error ", c, " time(s). Retry. *****"))

  }

  return(result)

}

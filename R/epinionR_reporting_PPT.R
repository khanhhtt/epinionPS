# ******************************************************************************
# Created: 9-Aug-2022
# Last modified: 28-Jun-2024
# Contributor(s): httk
# ******************************************************************************
# May you do good and not evil.
# May you find forgiveness for yourself and forgive others.
# May you share freely, never taking more than you give.
# ******************************************************************************

options(error = recover )

#' Generate Powerpoint report
#'
#' This function allows users to generate a simple Epinion Standard Powerpoint report based on
#' a prepared Powerpoint template and an OrderForm.
#'
#' @param df datainput (dataframe).
#' @param orderform_folder directory where orderforms stored. Default is the current working dir.
#' @param output_folder directory where output report should be placed. Default is the current working dir.
#' @param listOrderForm list of Orderform.
#' @param ... Additional arguments may apply.
#' @return Epinion Standard Powerpoint report
#' @examples
#' # Set the example folder of the package as the working directory
#' setwd(system.file("extdata", package = "epinionPS"))
#'
#' # Read the data
#' mydata <- epinion_read_data(file = example("sample_test_report.sav"))
#'
#' # State the default template to generate report
#' report_template <- "Template_New.pptx"
#'
#' # Get all PPT OrderForms in the folder
#' list_orderForm <- list.files(path = paste0(getwd(), "/"), pattern='OrderForm_Graphic_*')
#'
#' # Run reports
#' epinion_reporting_PPT(mydata,
#'                       orderform_folder = paste0(getwd(), "/"),
#'                       output_folder = paste0(getwd(), "/"),
#'                       list_orderForm)

epinion_reporting_PPT = function(df,
                                 orderform_folder = getwd(),
                                 output_folder = getwd(),
                                 listOrderForm) {

  if (!require(pacman)) install.packages("pacman")
  pacman::p_load(tidyverse,
                 tools,
                 devtools,
                 anesrake,
                 haven,
                 openxlsx,
                 expss,
                 labelled,
                 lubridate,
                 janitor,
                 officer,
                 crosstable,
                 flextable,
                 mschart)

  if (any(grepl("~$OrderForm_Graphic", listOrderForm, fixed = TRUE)) == TRUE) {
    stop("The orderform is being opened by Excel application. Please close the file before running the syntax.")
  }

  for (orderForm in listOrderForm) {
    # df <- haven::as_factor(df, only_labelled = TRUE)
    if (!file.exists(paste0(orderform_folder, orderForm))) {
      stop(paste0("'", paste0(orderform_folder, orderForm), "'", " does not exist in the working folder."))

    } else {
      start_time <- Sys.time()

      epinion_create_PPT_report_main(df, orderform_folder = orderform_folder,
                                     output_folder = output_folder,
                                     report_orderform = orderForm)

      end_time <- Sys.time()
      print(end_time - start_time)

    }

  }

}

epinion_create_PPT_report_main = function(df,
                                          orderform_folder,
                                          output_folder,
                                          report_orderform) {

  # 01. Read information in the order form
  ## Sheet "GeneralInformation"
  report_generalInfor <- openxlsx::readWorkbook(paste0(orderform_folder, report_orderform),
                                                sheet = "GeneralInformation",
                                                colNames = TRUE,
                                                rowNames = FALSE,
                                                skipEmptyRows = FALSE,
                                                check.names = FALSE,
                                                na.strings = ""
                                                )

  report_output <- paste0(output_folder, report_generalInfor[1, 2], ".pptx")
  client_name <- report_generalInfor[2, 2]
  report_title <- report_generalInfor[4, 2]
  report_date <- format(Sys.Date(), "%d.%b %Y")

  if (file.exists(report_output) ) {
    out_temp = try(file(report_output, open = "w"),
                   silent = TRUE)

    if (suppressWarnings("try-error" %in% class(out_temp))) {
      stop(paste0("The output PPT report '", report_output,  "' is being opened by PowerPoint application. Please close the file before running the syntax."))
    }

    close(out_temp)
  }

  ## Sheet "Analysis"
  report_analysis <- openxlsx::readWorkbook(paste0(orderform_folder, report_orderform),
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

  section_slidenum <- list()

  for (i in 1:dim(section)[1]) {

    ### Add section slide
    if (i==1) {
      report_output <- epinion_addslide_section(template = report_template,
                                                section_title = section$Section[i],
                                                target = report_output)

      slidenum_tmp <- read_pptx(report_output) %>%
        length(.)-1
      section_slidenum <- append(section_slidenum, slidenum_tmp)

    } else {
      report_output <- epinion_addslide_section(template = report_output,
                                                section_title = section$Section[i],
                                                target = report_output)

      slidenum_tmp <- read_pptx(report_output) %>%
        length(.)-1
      section_slidenum <- append(section_slidenum, slidenum_tmp)

    }

    ### Add slides in each section

    pptx_obj <- read_pptx(report_output)

    for (j in section$start[i]:section$end[i]) {
      ## Get the common information
      # weight var
      if (report_generalInfor[5, 2] == "No") {
        weight_var = "totalt"

      } else if (report_generalInfor[5, 2] == "Yes") {
        weight_var = report_generalInfor[6, 2]
      }

      # weighted base: to determine if we would like to report weighted or unweighted base
      if (report_generalInfor[7, 2] == "Yes") {
        base = "base_chart"

      } else {
        base = "base_chart_unweighted"
      }

      # theme color
      if (is.na(report_analysis$Theme[j])) {
        theme_color = theme_default
      } else {
        theme_color = eval(parse(text=report_analysis$Theme[j]))
      }

      # object code
      if (grepl("COL",report_analysis$`Object.(CODE)`[j], fixed = TRUE)) {
        chart_direction = "vertical"
        label_direction = FALSE
      } else if (grepl("BAR",report_analysis$`Object.(CODE)`[j], fixed = TRUE)) {
        chart_direction = "horizontal"
        label_direction = TRUE
      } else if (report_analysis$`Object.(CODE)`[j] == "LINE01P") {
        if (class(report_analysis[["Independent.variables"]][j]) == "Date") {
          number_format = "yyyy-mm-dd"
        } else
          number_format = "General"
      }

      # layout
      layout_default <- "5_1 Content Placeholder"

      # filter
      if (is.na(report_analysis$`Filter/Condition.(If.requried)`[j])) {
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
      } else{
        sort_by_cat <- 'default'
      }

      if (!is.na(report_analysis$Sort_order[j]) & report_analysis$Sort_order[j] == "Ascending") {
        sort_order <- FALSE
      } else if (!is.na(report_analysis$Sort_order[j]) & report_analysis$Sort_order[j] == "Descending") {
        sort_order <- TRUE
      } else {
        sort_order <- report_analysis$Sort_order[j]
      }

      # base mode
      if (is.na(report_analysis$Base_mode[j])) {
        base_mode <- "Chart"
      } else {
        base_mode <- report_analysis$Base_mode[j]
      }

      # Add total bar
      if (is.na(report_analysis$Total_bar[j])) {
        total_bar = "default"
      } else {
        total_bar = report_analysis$Total_bar[j]
      }

      # Sigtest
      if (is.na(report_analysis$Sig_test[j])) {
        sigtest = FALSE
      } else {
        sigtest = TRUE
      }

      # significance level
      if (!is.na(report_analysis$Sig_level[j])){
        sig_level <- report_analysis$Sig_level[j]
      } else {
        sig_level <- 0.05
      }

      # Adjust bonferroni
      if (!is.na(report_analysis$Adjust_bonferroni[j]) & report_analysis$Adjust_bonferroni[j] == "Yes"){
        bonferroni <- TRUE
      } else {
        bonferroni <- FALSE
      }

      # Slide title, slide subtitle, chart title
      slide_title = ifelse(is.na(report_analysis$Slide.title[j]), "", report_analysis$Slide.title[j])
      slide_subtitle = ifelse(is.na(report_analysis$Slide.sub.title[j]), "",report_analysis$Slide.sub.title[j])
      chart_title = ifelse(is.na(report_analysis$Chart.Title[j]), "", report_analysis$Chart.Title[j])
      rp_analysis = report_analysis$Analysis[j]
      footer = ifelse(is.na(report_analysis$FootNote[j]), "", report_analysis$FootNote[j])

      ## Frequency of single variable
      if (rp_analysis == "Freq" & is.na(report_analysis$Advanced.Option[j])) {
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_freq_single(df, col_var, row_var, weight_var,
                                       slide_title, chart_title, slide_subtitle, footer,
                                       label_direction, chart_direction,
                                       theme_color, base,
                                       filter = filter,
                                       base_mode = base_mode,
                                       pptx_obj = pptx_obj,
                                       rp_analysis = rp_analysis,
                                       layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      } else if (rp_analysis == "Freq" & is.na(report_analysis$Advanced.Option[j]) == FALSE &
                 (report_analysis$`Object.(CODE)`[j] == "BAR02P" |
                  report_analysis$`Object.(CODE)`[j] == "COL01P")) {
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_freq_singles(df, col_var, row_var, weight_var,
                                        slide_title, chart_title, slide_subtitle, footer,
                                        label_direction, chart_direction,
                                        theme_color, base,
                                        filter = filter,
                                        base_mode = base_mode,
                                        sort_by_cat = sort_by_cat, sort_order = sort_order,
                                        pptx_obj = pptx_obj,
                                        rp_analysis = rp_analysis,
                                        layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Frequency of multiple variables
      if (rp_analysis == "Multiple Freq (By Case)") {
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_freq_multiple(df, col_var, row_var, weight_var,
                                         slide_title, chart_title, slide_subtitle, footer,
                                         label_direction, chart_direction,
                                         theme_color, base,
                                         counted_value = counted_value,
                                         filter = filter,
                                         base_mode = base_mode, total_bar = total_bar,
                                         sort_by_cat = sort_by_cat, sort_order = sort_order,
                                         pptx_obj = pptx_obj,
                                         rp_analysis = rp_analysis,
                                         layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)
      }

      ## Frequency of grid variables
      if (rp_analysis == "Freq" & is.na(report_analysis$Advanced.Option[j]) == FALSE &
          (report_analysis$`Object.(CODE)`[j] == "BAR04P" |
           report_analysis$`Object.(CODE)`[j] == "COL04P")) {
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_freq_grid(df, col_var, row_var, weight_var,
                                     slide_title, chart_title, slide_subtitle, footer,
                                     label_direction, chart_direction,
                                     theme_color, base,
                                     filter = filter,
                                     base_mode = base_mode,
                                     sort_by_cat = sort_by_cat, sort_order = sort_order,
                                     pptx_obj = pptx_obj,
                                     rp_analysis = rp_analysis,
                                     layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)
      }

      ## Frequency of grid variables with mean
      if (rp_analysis == "Freq Grid with Mean" & is.na(report_analysis$Advanced.Option[j]) == FALSE) {
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_freq_grid_mean(df, row_var, weight_var,
                                          slide_title, chart_title, slide_subtitle, footer,
                                          theme_color, base,
                                          filter = filter,
                                          base_mode = base_mode,
                                          sort_by_cat = sort_by_cat, sort_order = sort_order,
                                          pptx_obj = pptx_obj,
                                          rp_analysis = rp_analysis,
                                          layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)
      }

      ## Mean chart of single variables
      if (rp_analysis == "Mean chart") {
        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_meanbarchart_singles(df, col_var, row_var, weight_var,
                                                slide_title, chart_title, slide_subtitle, footer,
                                                label_direction, chart_direction,
                                                theme_color, base,
                                                filter = filter,
                                                base_mode = base_mode,
                                                # sort_by_cat = sort_by_cat,
                                                sort_order = sort_order,
                                                pptx_obj = pptx_obj,
                                                rp_analysis = rp_analysis,
                                                layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)
      }

      ## Crosstab mean: Mean chart of single variable
      if (rp_analysis == "Cross Mean") {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_meanchart_crosstab(df, col_var, row_var, weight_var,
                                              slide_title, chart_title, slide_subtitle, footer,
                                              label_direction, chart_direction,
                                              theme_color, base,
                                              counted_value = counted_value,
                                              filter = filter,
                                              base_mode = base_mode, total_bar = total_bar,
                                              sort_by_cat = sort_by_cat, sort_order = sort_order,
                                              sigtest = sigtest,
                                              sig_level = sig_level,
                                              bonferroni = bonferroni,
                                              pptx_obj = pptx_obj,
                                              rp_analysis = rp_analysis,
                                              layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      }


      ## Crosstab: single vs. single
      if (rp_analysis == "Crosstabs" &
          (grepl("BAR",report_analysis$`Object.(CODE)`[j], fixed = TRUE) |
           grepl("COL",report_analysis$`Object.(CODE)`[j], fixed = TRUE))) {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_crosstab_single(df, col_var, row_var, weight_var,
                                           slide_title, chart_title, slide_subtitle, footer,
                                           label_direction, chart_direction,
                                           theme_color, base,
                                           legend_pos = "b",
                                           filter = filter,
                                           base_mode = base_mode, total_bar = total_bar,
                                           sort_by_cat = sort_by_cat, sort_order = sort_order,
                                           sigtest = sigtest,
                                           sig_level = sig_level,
                                           bonferroni = bonferroni,
                                           pptx_obj = pptx_obj,
                                           rp_analysis = rp_analysis,
                                           layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Crosstab: single vs. multiple
      if (rp_analysis == "Multiple Crosstabs (By Case)") {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_crosstab_multiple(df, col_var, row_var, weight_var,
                                             slide_title, chart_title, slide_subtitle, footer,
                                             label_direction, chart_direction,
                                             theme_color, base,
                                             legend_pos = "b",
                                             counted_value = counted_value,
                                             filter = filter,
                                             base_mode = base_mode,
                                             total_bar = total_bar,
                                             sort_by_cat = sort_by_cat, sort_order = sort_order,
                                             sigtest = sigtest,
                                             sig_level = sig_level,
                                             bonferroni = bonferroni,
                                             pptx_obj = pptx_obj,
                                             rp_analysis = rp_analysis,
                                             layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      }


      if (grepl("Multiple Crosstabs Row (By Case)", rp_analysis, fixed = TRUE)) {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_crosstab_multiple(df, col_var, row_var, weight_var,
                                             slide_title, chart_title, slide_subtitle, footer,
                                             label_direction, chart_direction,
                                             theme_color, base,
                                             legend_pos = "b",
                                             counted_value = counted_value,
                                             filter = filter,
                                             base_mode = base_mode,
                                             total_bar = total_bar,
                                             sort_by_cat = sort_by_cat, sort_order = sort_order,
                                             sigtest = sigtest,
                                             sig_level = sig_level,
                                             bonferroni = bonferroni,
                                             pptx_obj = pptx_obj,
                                             rp_analysis = rp_analysis,
                                             layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

	  }


      ## Crosstab: single vs. grid
      if (rp_analysis == "Crosstabs Row" | rp_analysis == "Crosstabs Row with Mean") {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_crosstab_grid(df, col_var, row_var, weight_var,
                                         slide_title, chart_title, slide_subtitle, footer,
                                         label_direction, chart_direction,
                                         theme_color, base,
                                         legend_pos = "b",
                                         filter = filter,
                                         base_mode = base_mode,
                                         total_bar = total_bar,
                                         sort_by_cat = sort_by_cat, sort_order = sort_order,
                                         sigtest = sigtest,
                                         sig_level = sig_level,
                                         bonferroni = bonferroni,
                                         pptx_obj = pptx_obj,
                                         rp_analysis = rp_analysis,
                                         layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      }


      ## Line chart
      if (report_analysis$`Object.(CODE)`[j] == "LINE01N" &
          is.na(report_analysis$Advanced.Option[j]) == TRUE) {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]
        if (is.na(report_analysis$Group_by[j])) {
          group_var = "totalt"
          legend_position = "n"
        } else {
          group_var = report_analysis$Group_by[j]
          legend_position = "b"
        }

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_linechart_cro(df, col_var, row_var, weight_var, group_var,
                                         slide_title, chart_title, slide_subtitle, footer,
                                         theme_color, base,
                                         legend_pos = legend_position,
                                         filter = filter,
                                         counted_value = counted_value,
                                         total_line = total_bar,
                                         sigtest = sigtest,
                                         sig_level = sig_level,
                                         bonferroni = bonferroni,
                                         pptx_obj = pptx_obj,
                                         rp_analysis = rp_analysis,
                                         layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      }

      if (report_analysis$`Object.(CODE)`[j] == "LINE01N" &
          is.na(report_analysis$Advanced.Option[j]) == FALSE) {
        col_var = report_analysis[["Independent.variables"]][j]
        row_var = report_analysis[["Dependent.variables"]][j]
        if (is.na(report_analysis$Group_by[j]) == FALSE) {
          stop("Could not group grid variables on 1 table. Please update the orderform!")
        }

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_linechart_cro_grid(df, col_var, row_var, weight_var,
                                              slide_title, chart_title, slide_subtitle, footer,
                                              theme_color, base,
                                              legend_pos = "b",
                                              filter = filter,
                                              counted_value = counted_value,
                                              total_line = total_bar,
                                              sigtest = sigtest,
                                              sig_level = sig_level,
                                              bonferroni = bonferroni,
                                              pptx_obj = pptx_obj,
                                              rp_analysis = rp_analysis,
                                              layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      }

      ## Open answer
      if (rp_analysis == "Open answer") {
        report_mode = "unique"

        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]
        layout = "6_1 Content Placeholder"

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_openans(df, row_var,
                                   slide_title, chart_title, slide_subtitle, footer,
                                   filter = filter,
                                   report_mode = report_mode,
                                   pptx_obj = pptx_obj,
                                   rp_analysis = rp_analysis,
                                   layout = layout_default)
        )

        print_status(rp_analysis, row_var, col_var)

      } else if (rp_analysis == "Open answer - Full") {
        report_mode = "full"

        col_var = "totalt"
        row_var = report_analysis[["Independent.variables"]][j]
        layout = "6_1 Content Placeholder"

        pptx_obj <- catchErrorAndRetry(
          epinion_addslide_openans(df, row_var,
                                   slide_title, chart_title, slide_subtitle, footer,
                                   filter = filter,
                                   report_mode = report_mode,
                                   pptx_obj = pptx_obj,
                                   rp_analysis = rp_analysis,
                                   layout = layout_default)
          )

        print_status(rp_analysis, row_var, col_var)
      }


    }
    print(paste0("================ Please wait for section '", section$Section[i] ,"' to be saved ================"))
    print(pptx_obj, target = report_output)
  }

  print("================ Please wait for the report to be saved ʢ´• ᴥ •`ʡ ================")

  # Update slide table of content and cover slide
  section$section_slidenum <- do.call(rbind, section_slidenum)
  TOC_pptx <- epinion_format_table_PPT(section[,c("Section", "section_slidenum")], font_size = 14, font_color = "#e14646") %>%
    height_all(height = 0.47, part = "all") %>%
    width(j = c(1, 2), width = c(6.61, 0.59)) %>%
    align(j = c(1, 2), align = c("left", "center"), part = "all") %>%
    delete_part(part = "header")

  pptx_obj <- read_pptx(report_output)
  ppt <- pptx_obj %>%
    on_slide(index = 1) %>%
    ph_with(value = report_title, location = ph_location_label(ph_label = "Report Title")) %>%
    ph_with(value = client_name, location = ph_location_label(ph_label = "Client Name")) %>%
    ph_with(value = report_date, location = ph_location_label(ph_label = "Report Date")) %>%
    on_slide(index = 2) %>%
    ph_with(value = TOC_pptx, location = ph_location_label(ph_label = "Content Placeholder 3")) %>%
    print(report_output)

}

# ##############################################################################
# Section slide
# ##############################################################################
# Add section slide to PPT report
epinion_addslide_section = function(template,
                                    section_title = empty_content(),
                                    section_subtitle = empty_content(),
                                    target) {

  pptx_obj <- read_pptx(template) %>%
    add_slide(layout = "3_Section", master = "Office-tema") %>%
    ph_with(value = section_title, location = ph_location_label(ph_label = "Section Title")) %>%
    ph_with(value = section_subtitle, location = ph_location_label(ph_label = "Text Placeholder")) %>%
    move_slide(index = length(.), to = length(.) - 1)%>%
    print(target = target)

}

# ##############################################################################
# Prepare input calculation
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
      mutate(percentage = Total/100,
             base_cat_label = paste0(row_labels, " (N=", round_half_up(`Total|Count`,0), ")"),
             base_cat_label_unweighted = paste0(row_labels, " (N=", round_half_up(`Total|Count_unweight`,0), ")"),
             base_chart = round_half_up(sum(`Total|Count`, na.rm = TRUE)/2, 0),
             base_chart_unweighted = round_half_up(sum(`Total|Count_unweight`, na.rm = TRUE)/2, 0),
             totalt = 1) %>%
      filter(row_labels != "#Total")

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
      arrange(row_varOrder) %>%
      mutate(base_chart = round_half_up(`#Total|Count`, 0),
             base_chart_unweighted = round_half_up(`#Total|Count_unweight`, 0),
             base_cat_label = paste0(row_labels, " (N=", base_chart, ")"),
             base_cat_label_unweighted = paste0(row_labels, " (N=", base_chart_unweighted, ")"))

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
      mutate(Total = ifelse(row_labels == "#Total", sum(Total, na.rm = TRUE)-100, Total),
             base_chart = expss::index_col(expss::match_col("#Total", row_labels), `Total|Count`),
             base_chart_unweighted = expss::index_col(expss::match_col("#Total", row_labels), `Total|Count_unweight`)) %>%
      filter(row_labels != "#Total") %>%
      adorn_totals("row", fill = NA_integer_, na.rm = TRUE, name = "Total", Total, `Total|Count`, `Total|Count_unweight`) %>%
      mutate(base_chart = round_half_up(max(base_chart, na.rm = TRUE), 0),
             base_chart_unweighted = round_half_up(max(base_chart_unweighted, na.rm = TRUE), 0),
             percentage = Total/100,
             base_cat_label = paste0(row_labels, " (N=", round_half_up(`Total|Count`,0), ")"),
             base_cat_label_unweighted = paste0(row_labels, " (N=", round_half_up(`Total|Count_unweight`,0), ")"),
             totalt = 1)

  }

  df_tbl

}

# Prepare mean table
epinion_calc_mean = function(x, col_var, row_var, weight_var,
                             filter_var, filter_val,
                             rp_analysis) {

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
    tab_weight(weight = weight_value)

  if (grepl("with Mean", rp_analysis, fixed = TRUE) > 0) {
    df_tbl <- df_tbl %>%
      tab_stat_mean(label = "Snitt") %>%
      tab_pivot() %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      separate(row_labels, c("row_varOrder", "row_labels"), sep = "\\|", remove = TRUE) %>%
      mutate(Snitt = round_half_up(Snitt, 2)) %>%
      select(., -row_varOrder)

  } else if (rp_analysis == "Mean chart") {
    df_tbl <- df_tbl %>%
      tab_stat_mean_sd_n(weighted_valid_n = TRUE) %>%
      tab_weight(weight = total()) %>%
      tab_stat_mean_sd_n(weighted_valid_n = FALSE,
                         labels = c("Mean_unweight", "Std. dev. unweight", "Unw. valid N")) %>%
      tab_pivot() %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      separate(row_labels, c("row_varOrder", "row_labels"), sep = "\\|", remove = TRUE) %>%
      select(., -row_varOrder) %>%
      mutate(base_chart = round_half_up(max(`Valid N`, na.rm = TRUE),0),
             base_chart_unweighted = round_half_up(max(`Unw. valid N`, na.rm = TRUE),0),
             base_cat_label = paste0(row_labels, " (N=", trimws(round_half_up(`Valid N`, 0)), ")"),
             base_cat_label_unweighted = paste0(row_labels, " (N=", trimws(round_half_up(`Unw. valid N`, 0)), ")"),
             totalt = 1)
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
    if_na(0)
    # as.data.frame(.)

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
      tab_stat_cpct() %>%
      tab_last_sig_cpct(sig_level = sig_level, bonferroni = bonferroni, digits = 0,
                        subtable_marks = "greater", sig_labels = LETTERS) %>%
      tab_pivot() %>%
      epinion_add_percent(excluded_rows = "#") %>%
      as.data.frame(.) %>%
      pivot_longer(., !row_labels, names_to = "var", values_to = "sigPct") %>%
      separate(var, c("value", "sigLetter"), sep = "\\|", remove = TRUE) %>%
      filter(., row_labels != "#Total cases")

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

  if (length(col_var) == 1 & length(row_var) == 1) {
    df_tbl_input <- df_tbl_input %>%
      filter(!is.na(!!as.symbol(col_var))) %>%
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
    df_tbl_input <- df_tbl_input %>%
      filter(!is.na(!!as.symbol(col_var))) %>%
      tab_cells(..$head(row_var, 1) %to% tail(row_var, 1))  %>%
      tab_cols("|" = unvr(..$col_var ), total(label = "Total"))

  } else if (length(col_var) > 1 & length(row_var) > 1) {
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
      select(., totalt, !!as.symbol(weight_var), all_of(row_var), all_of(col_var)) %>%
      filter(if_any(head(col_var, 1):tail(col_var, 1), ~. == 1)) %>%
      tab_cells(..$head(row_var, 1) %to% tail(row_var, 1))  %>%
      tab_cols(mdset(..$head(col_var, 1) %to% ..$tail(col_var, 1)), total(label = "Total"))

  }

  if (length(col_var) >= 1 & length(row_var) == 1) {
    df_tbl <- df_tbl_input %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_mean_sd_n(weighted_valid_n = TRUE,
                         labels = c("Snitt", "Std. dev.", "Valid N")) %>%
      tab_weight(weight = total()) %>%
      tab_stat_mean_sd_n(weighted_valid_n = FALSE,
                         labels = c("Mean_unweight", "Std. dev. unweight", "Unw. valid N")) %>%
      tab_pivot() %>%
      as.data.frame(.)

  } else if (length(col_var) >= 1 & length(row_var) > 1) {
    df_tbl <- df_tbl_input %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_mean_sd_n(weighted_valid_n = TRUE) %>%
      tab_weight(weight = total()) %>%
      tab_stat_mean_sd_n(weighted_valid_n = FALSE,
                         labels = c("Mean_unweight", "Std. dev. unweight", "Unw. valid N")) %>%
      tab_pivot() %>%
      as.data.frame(.)

  }

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
      filter(grepl("Mean", row_labels)) %>%
      pivot_longer(., -row_labels, names_to = "var", values_to = "Snitt_sigtest") %>%
      separate(var, c("column_var", "sigLetter"), sep = "\\|", remove = TRUE) %>%
      mutate(row_labels = gsub("\\|Mean", "", row_labels))

  }

  out

}


# ##############################################################################
# Bar chart
# ##############################################################################
# Formating bar chart
epinion_barchart_formating = function (chart_obj, colour_list,
                                       chart_direction = "vertical",
                                       grouping = "clustered",
                                       label_direction = FALSE
                                       ) {

  if (grouping == "stacked" | grouping == "percentStacked") {
    chart_data_labels_position = "ctr"
    color_labels = "#ffffff"
  } else {
    chart_data_labels_position = "outEnd"
    color_labels = "#0f283c"
  }

  chart_obj <- chart_obj %>%
    chart_settings(dir = chart_direction, # dir = "vertical" for column chart, dir = "horizontal" for bar chart
                   gap_width = 150,
                   overlap = -30,
                   grouping = grouping) %>%
    chart_theme(legend_text = fp_text(color = "#0f283c", font.size = 10),
                axis_text_x = fp_text(color = "#0f283c", font.size = 10),
                axis_text_y = fp_text(color = "#0f283c", font.size = 10),
                axis_ticks_x = fp_border(color = "#8a8d91", width = 0.5),
                axis_ticks_y = fp_border(color = "#8a8d91", width = 0.5),
                grid_major_line_x = fp_border(width = 0),
                grid_major_line_y = fp_border(width = 0)) %>%
    chart_ax_x(major_tick_mark = 'none') %>%
    chart_ax_y(display = 0,
               limit_min = 0,
               limit_max = 1,
               num_fmt = "0%%") %>%
    # enable chart data labels
    chart_data_labels(show_val = TRUE,
                      position = chart_data_labels_position,
                      num_fmt = "0%") %>%
    # hide chart title
    chart_labels()

  series_names <- mschart:::get_series_names(chart_obj)

  if (length(colour_list) > 1) {
    if(length(series_names) <= length(colour_list)) {
      palette_ <- colour_list[[length(series_names)]][order(1:length(series_names), decreasing = label_direction)]
    } else {
      palette_ <- sample(colors(), size = length(series_names), replace = TRUE)
    }
  } else {
    palette_ <- colour_list[[1]]
  }



  chart_obj$series_settings <- list(
    # chart_data_fill: color of series data fill in PPT
    fill = setNames(palette_, series_names),
    # chart_data_stroke: border's color of series data in PPT
    colour = setNames(palette_, series_names),
    # chart_data_line_width: border's width of series data in PPT
    line_width = setNames(rep(0, length(series_names)), series_names),
    # chart_labels_text: legend of the chart in PPT
    labels_fp = setNames(rep(list(fp_text(color = color_labels, font.size = 10)), length(series_names)), series_names)
  )

  chart_obj
}

# ==============================================================================
# Add frequency chart of single variable to PPT
epinion_addslide_freq_single = function(x, col_var, row_var, weight_var,
                                        slide_title, chart_title, slide_subtitle, footer,
                                        label_direction = FALSE, chart_direction,
                                        theme_color, base, legend_pos = "n",
                                        filter = "totalt=1", base_mode = "Chart",
                                        pptx_obj, rp_analysis, layout) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_chart <- epinion_calc_freq(x, col_var, row_var, weight_var,
                               filter_var, filter_val,
                               rp_analysis)

  if (base_mode == "Chart") {
    row_var_chart = "row_labels"
  } else if (base == "base_chart") {
    row_var_chart = "base_cat_label"
  } else if (base == "base_chart_unweighted") {
    row_var_chart = "base_cat_label_unweighted"
  }

  df_chart[[row_var_chart]] <- factor(df_chart[[row_var_chart]], levels = unique(df_chart[[row_var_chart]])[order(1:length(unique(df_chart[[row_var_chart]])), decreasing = label_direction)])
  df_chart[[col_var]] <- factor(df_chart[[col_var]], levels = unique(df_chart[[col_var]])[order(1:length(unique(df_chart[[col_var]])), decreasing = label_direction)])


  chart <-  ms_barchart(data = df_chart,
                        x = row_var_chart,
                        y = "percentage",
                        group = col_var)

  chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                      chart_direction = chart_direction,
                                      label_direction = label_direction) %>%
    chart_theme(legend_position = legend_pos)%>%
    chart_ax_y(display = 0,
               limit_min = 0,
               limit_max = ifelse(max(df_chart$percentage) < 1, 1, max(df_chart$percentage) + 0.2),
               num_fmt = "0%%")

  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(df_chart[[base]]))),
                               layout = layout
  )

  pptx_obj

}

# ==============================================================================
# Add frequency chart of single variables (same categories) to PPT
epinion_addslide_freq_singles = function(x, col_var, row_var, weight_var,
                                         slide_title, chart_title, slide_subtitle, footer,
                                         label_direction = FALSE, chart_direction,
                                         theme_color, base,
                                         filter = "totalt=1", base_mode = "Chart",
                                         sort_by_cat = 'default', sort_order = TRUE,
                                         pptx_obj, rp_analysis, layout) {

  var_grid <- strsplit(row_var, "\\,")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_chart <- epinion_calc_freq(x, col_var, var_grid, weight_var,
                               filter_var, filter_val,
                               rp_analysis)

  # sort
  if (sort_by_cat == 'default') {
    df_chart <- df_chart

  } else {
    if (sort_order == TRUE) {
      df_chart <- df_chart %>%
        arrange(desc(!!as.symbol(sort_by_cat)))

    } else if (sort_order == FALSE) {
      df_chart <- df_chart %>%
        arrange(!!as.symbol(sort_by_cat))

    }
  }

  out <- df_chart %>%
    select(., row_labels, !contains("Count"), - row_varOrder) %>%
    select(., row_labels, !contains("base_"), -`#Total`) %>%
    mutate(across(where(is.numeric), ~./100)) %>%
    pivot_longer(., !row_labels, names_to = "value", values_to = "percentage") %>%
    left_join(., select(df_chart, row_labels, base_chart, base_chart_unweighted, base_cat_label, base_cat_label_unweighted),
              by = "row_labels")

  if (base_mode == "Chart") {
    row_var_chart = "row_labels"
  } else if (base == "base_chart") {
    row_var_chart = "base_cat_label"
  } else if (base == "base_chart_unweighted") {
    row_var_chart = "base_cat_label_unweighted"
  }


  out[[row_var_chart]] <- factor(out[[row_var_chart]], levels = unique(df_chart[[row_var_chart]])[order(1:length(unique(df_chart[[row_var_chart]])), decreasing = label_direction)])
  out[["value"]] <- factor(out[["value"]], levels = unique(out[["value"]])[order(1:length(unique(out[["value"]])), decreasing = label_direction)])

  chart <-  ms_barchart(data = out,
                        x = "value",
                        y = "percentage",
                        group = row_var_chart)

  chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                      label_direction = label_direction,
                                      chart_direction = chart_direction
  ) %>%
    chart_ax_y(display = 0,
               limit_min = 0,
               limit_max = ifelse(max(out$percentage) < 1, 1, max(out$percentage) + 0.2),
               num_fmt = "0%%")

  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = ifelse(base_mode == "Category", "",
                                                   ifelse(length(unique(round_half_up(out[[base]],0)))==1,
                                                          paste0("N=", unique(round_half_up(out[[base]],0))),
                                                          "")),
                               layout = layout)

}

# ==============================================================================
# Add frequency chart of grid variables to PPT
epinion_addslide_freq_grid = function(x, col_var, row_var, weight_var,
                                      slide_title, chart_title, slide_subtitle, footer,
                                      label_direction = FALSE, chart_direction,
                                      theme_color, base,
                                      filter = "totalt=1", base_mode = "Category",
                                      sort_by_cat = 'default', sort_order = TRUE,
                                      pptx_obj, rp_analysis, layout) {

  var_grid <- strsplit(row_var, "\\,")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_chart <- epinion_calc_freq(x, col_var, var_grid, weight_var,
                                filter_var, filter_val,
                                rp_analysis)

  # sort
  if (sort_by_cat == 'default') {
    df_chart <- df_chart

  } else {
    if (sort_order == TRUE) {
      df_chart <- df_chart %>%
        arrange(desc(!!as.symbol(sort_by_cat)))

    } else if (sort_order == FALSE) {
      df_chart <- df_chart %>%
        arrange(!!as.symbol(sort_by_cat))

    }
  }

  out <- df_chart %>%
    select(., row_labels, !contains("Count"), - row_varOrder) %>%
    select(., row_labels, !contains("base_"), -`#Total`) %>%
    mutate(across(where(is.numeric), ~./100)) %>%
    pivot_longer(., !row_labels, names_to = "value", values_to = "percentage") %>%
    left_join(., select(df_chart, row_labels, base_chart, base_chart_unweighted, base_cat_label, base_cat_label_unweighted),
              by = "row_labels")

  if (base_mode == "Chart") {
    row_var_chart = "row_labels"
  } else if (base == "base_chart") {
    row_var_chart = "base_cat_label"
  } else if (base == "base_chart_unweighted") {
    row_var_chart = "base_cat_label_unweighted"
  }

  out[[row_var_chart]] <- factor(out[[row_var_chart]], levels = unique(df_chart[[row_var_chart]])[order(1:length(unique(df_chart[[row_var_chart]])), decreasing = label_direction)])
  out[["value"]] <- factor(out[["value"]], levels = levels(as_factor(x[[var_grid[1]]])))

  chart <-  ms_barchart(data = out,
                        x = row_var_chart,
                        y = "percentage",
                        group = "value")

  chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                      label_direction = label_direction,
                                      grouping = "stacked") %>%
    epinion_as_bar_stack(dir = chart_direction, gap_width = 100)

  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               # base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(round_half_up(df_chart[[base]]),0))),
                               base_chart = "",
                               layout = layout)

}

# ==============================================================================
# Add frequency chart with mean table of grid variables to PPT
epinion_addslide_freq_grid_mean = function(x, row_var, weight_var,
                                           slide_title, chart_title, slide_subtitle, footer,
                                           theme_color, base,
                                           filter = "totalt=1", base_mode = "Chart",
                                           sort_by_cat = 'default', sort_order = TRUE,
                                           pptx_obj, rp_analysis, layout) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  var_grid <- strsplit(row_var, "\\,")[[1]]

  # Prepare mean dataframe
  var_grid_mean <- paste0(var_grid, "_mean")

  df_chart_mean <- epinion_calc_mean(x, col_var, var_grid_mean, weight_var,
                                    filter_var, filter_val,
                                    rp_analysis)


  # Prepare percentage dataframe
  df_chart <- epinion_calc_freq(x, col_var, var_grid, weight_var,
                                filter_var, filter_val,
                                rp_analysis) %>%
    left_join(., df_chart_mean, by = "row_labels")

  # sort
  if (sort_by_cat == 'default') {
    df_chart <- df_chart

  } else {
    if (sort_order == TRUE) {
      df_chart <- df_chart %>%
        arrange(desc(!!as.symbol(sort_by_cat)))

    } else if (sort_order == FALSE) {
      df_chart <- df_chart %>%
        arrange(!!as.symbol(sort_by_cat))

    }
  }

  out <- df_chart %>%
    select(., row_labels, !contains("Count"), - row_varOrder) %>%
    select(., row_labels, !contains("base_"), -`#Total`) %>%
    mutate(across(where(is.numeric), ~./100)) %>%
    pivot_longer(., !row_labels, names_to = "value", values_to = "percentage") %>%
    left_join(., select(df_chart, row_labels, base_chart, base_chart_unweighted, base_cat_label, base_cat_label_unweighted, Snitt),
              by = "row_labels")

  if (base_mode == "Chart") {
    row_var_chart = "row_labels"
  } else if (base == "base_chart") {
    row_var_chart = "base_cat_label"
  } else if (base == "base_chart_unweighted") {
    row_var_chart = "base_cat_label_unweighted"
  }

  out[[row_var_chart]] <- factor(out[[row_var_chart]], levels = unique(df_chart[[row_var_chart]])[order(1:length(unique(df_chart[[row_var_chart]])), decreasing = TRUE)])
  out[["value"]] <- factor(out[["value"]], levels = levels(as_factor(x[[var_grid[1]]])))

  # Create chart
  chart <-  ms_barchart(data = out,
                        x = row_var_chart,
                        y = "percentage",
                        group = "value")

  chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                      grouping = "stacked") %>%
    epinion_as_bar_stack(dir = "horizontal", gap_width = 100)

  # Create mean table
  mean_table <- unique(out[,c("row_labels", "Snitt")]) %>%
    select(.,-row_labels)

  ft <- epinion_format_table_PPT(mean_table) %>%
    width(width = 1.35) %>%
    height_all(height = (4.15-0.23-0.48)/(dim(mean_table)[1])) %>%
    height(height=0.23, part = "header")
  # 4.15 is the height of Content place holder in the slide template
  # 0.48 is the height for legend of the chart

  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               # base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(round_half_up(out[[base]]),0))),
                               layout = layout,
                               mean_tbl_obj = ft)

}


# ==============================================================================
# Add frequency chart of multiple variables to PPT
epinion_addslide_freq_multiple = function(x, col_var, row_var, weight_var,
                                          slide_title, chart_title, slide_subtitle, footer,
                                          label_direction = FALSE, chart_direction,
                                          theme_color, base,
                                          legend_pos = "n", counted_value = "Yes",
                                          filter = "totalt=1", base_mode = "Chart", total_bar = "default",
                                          sort_by_cat = 'default', sort_order = TRUE,
                                          pptx_obj, rp_analysis, layout) {
  var_multiple <- strsplit(row_var, "\\+")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_chart <- epinion_calc_freq(x, col_var, var_multiple, weight_var,
                               filter_var, filter_val,
                               rp_analysis,
                               counted_value = counted_value)

  if (base_mode == "Chart") {
    row_var_chart = "row_labels"
  } else if (base == "base_chart") {
    row_var_chart = "base_cat_label"
  } else if (base == "base_chart_unweighted") {
    row_var_chart = "base_cat_label_unweighted"
  }

  # sort option
  if (sort_by_cat == 'default') {
    df_chart <- df_chart
  } else {
    excluded_cats <- c(strsplit(sort_by_cat, "\\,")[[1]], "Total")
    df_sort_0 <- filter(df_chart, row_labels %in% excluded_cats)
    df_sort_1 <- filter(df_chart, !row_labels %in% excluded_cats)
    df_sort_1 <- df_sort_1[order(df_sort_1$percentage, decreasing = sort_order), ]
    df_chart <- bind_rows(df_sort_1, df_sort_0)
  }

  # Add total bar
  if (total_bar == "default") {
    df_chart <- df_chart %>%
      filter(row_labels != "Total")
  } else if (total_bar == "At first") {
    suppressMessages(if (!require("collapse", quietly = TRUE)) install.packages("collapse", quiet = TRUE))
    df_chart <- collapse::roworderv(df_chart, neworder = df_chart$row_labels == "Total")
  } else if (total_bar == "At last") {
    # do nothing
  }

  df_chart[[row_var_chart]] <- factor(df_chart[[row_var_chart]], levels = unique(df_chart[[row_var_chart]])[order(1:length(unique(df_chart[[row_var_chart]])), decreasing = label_direction)])



  chart <-  ms_barchart(data = df_chart,
                        x = row_var_chart,
                        y = "percentage",
                        group = "totalt")

  chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                      chart_direction = chart_direction,
                                      label_direction = TRUE) %>%
    chart_theme(legend_position = legend_pos) %>%
    chart_ax_y(display = 0,
               limit_min = 0,
               limit_max = ifelse(max(df_chart$percentage) < 1, 1, max(df_chart$percentage) + 0.2),
               num_fmt = "0%%")


  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(round_half_up(df_chart[[base]]),0))),
                               layout = layout)

}


# ==============================================================================
# Add mean bar chart of single variables to PPT
epinion_addslide_meanbarchart_singles = function(x, col_var, row_var, weight_var,
                                                 slide_title, chart_title, slide_subtitle, footer,
                                                 label_direction = FALSE, chart_direction,
                                                 theme_color, base,
                                                 filter = "totalt=1", base_mode = "Chart",
                                                 sort_by_cat = 'default', sort_order = TRUE,
                                                 pptx_obj, rp_analysis, layout) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  var_grid <- strsplit(row_var, "\\,")[[1]]

  df_chart <- epinion_calc_mean(x, col_var, var_grid, weight_var,
                               filter_var, filter_val,
                               rp_analysis)

  if (base_mode == "Chart") {
    row_var_chart = "row_labels"
  } else if (base == "base_chart") {
    row_var_chart = "base_cat_label"
  } else if (base == "base_chart_unweighted") {
    row_var_chart = "base_cat_label_unweighted"
  }

  # sort option
  if (is.na(sort_order)) {
    df_sort <- df_chart
  } else {
    df_sort <- df_chart[order(df_chart$Mean, decreasing = sort_order), ]
  }

  df_chart[[row_var_chart]] <- factor(df_chart[[row_var_chart]], levels = df_sort[[row_var_chart]][order(1:length(df_sort[[row_var_chart]]), decreasing = label_direction)])

  chart <-  ms_barchart(data = df_chart,
                        x = row_var_chart,
                        y = "Mean",
                        group = col_var)

  chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                      chart_direction = chart_direction,
                                      label_direction = label_direction) %>%
    chart_theme(legend_position = "n") %>%
    chart_ax_y(display = 0,
               limit_min = 0,
               limit_max = ifelse(max(df_chart$Mean) < 5, 5, max(df_chart$Mean) + 1),
               num_fmt = "0.00") %>%
    chart_data_labels(show_val = TRUE,
                      position = "outEnd",
                      num_fmt = "0.00")

  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(df_chart[[base]]))),
                               layout = layout)

}


# ==============================================================================
# Add crosstab chart of single variable to PPT
epinion_addslide_crosstab_single = function(x, col_var, row_var, weight_var,
                                            slide_title, chart_title, slide_subtitle, footer,
                                            label_direction = FALSE, chart_direction,
                                            theme_color, base, legend_pos = "n",
                                            filter = "totalt=1", base_mode = "Chart", total_bar = "default",
                                            sort_by_cat = 'default', sort_order = TRUE, sigtest = FALSE,
                                            sig_level = 0.05,
                                            bonferroni = FALSE,
                                            pptx_obj, rp_analysis, layout) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_chart <- epinion_calc_crosstab(x, col_var, row_var, weight_var,
                                   filter_var, filter_val,
                                   rp_analysis) %>%
    as.data.frame(.) %>%
    mutate(base_chart = round_half_up(expss::index_col(expss::match_col("#Total", row_labels), `Total|Count`), 0),
           base_chart_unweighted = round_half_up(expss::index_col(expss::match_col("#Total", row_labels), `Total|Count_unweight`), 0),
          ) %>%
    filter(row_labels != "#Total")

  df_count <- df_chart %>%
    select(., row_labels, ends_with("Count")) %>%
    rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
    adorn_totals(., where = "row", fill = "-", na.rm = TRUE, name = "Total") %>%
    mutate(across(where(is.numeric), round_half_up)) %>%
    filter(row_labels == "Total") %>%
    pivot_longer(., -row_labels, names_to = "value", values_to = "count")

  df_count_unweight <- df_chart %>%
    select(., row_labels, ends_with("Count_unweight")) %>%
    rename_with(~ sub("\\|Count_unweight$", "", .x), everything()) %>%
    adorn_totals(., where = "row", fill = "-", na.rm = TRUE, name = "Total") %>%
    filter(row_labels == "Total") %>%
    pivot_longer(., -row_labels, names_to = "value", values_to = "count_unweight")

  # sort option
  if (sort_by_cat == 'default' & is.na(sort_order)) {
    # do nothing
  } else if (sort_by_cat == 'default' & !is.na(sort_order) & sort_order == TRUE) {
    # Sort desc by Total
    df_chart <- df_chart %>%
      arrange(desc(Total))

  } else if (sort_by_cat == 'default' & !is.na(sort_order) & sort_order == FALSE) {
    # Sort asc by Total
    df_chart <- df_chart %>%
      arrange(Total)
  }

  # Calculate sigtest
  ctable_sigtest <- epinion_calc_crosstab(x, col_var, row_var, weight_var,
                                          filter_var, filter_val,
                                          rp_analysis,
                                          sigtest = TRUE,
                                          sig_level = sig_level,
                                          bonferroni = bonferroni)


  out <- df_chart %>%
    select(., row_labels, !contains("Count")) %>%
    select(., row_labels, !contains("base_"), -"Total") %>%
    mutate(across(where(is.numeric), ~./100)) %>%
    pivot_longer(., !row_labels, names_to = "value", values_to = "percentage") %>%
    left_join(., select(df_count, -row_labels), by = "value") %>%
    left_join(., select(df_count_unweight, -row_labels), by = "value") %>%
    mutate(base_cat_label = paste0(value, " (N=", trimws(round_half_up(count, 0)), ")"),
           base_cat_label_unweighted = paste0(value, " (N=", trimws(round_half_up(count_unweight, 0)), ")")
          ) %>%
    left_join(., select(df_chart, row_labels, base_chart, base_chart_unweighted), by = "row_labels") %>%
    left_join(., ctable_sigtest, by = c("row_labels", "value")) %>%
    mutate(col_var_label = paste0(value, " \n (", sigLetter, ")"),
           col_var_label_with_base = paste0(base_cat_label, " \n (", sigLetter, ")"),
           col_var_label_with_base_unweighted = paste0(base_cat_label_unweighted, " \n (", sigLetter, ")"))

  if (!sigtest) {
    if (base_mode == "Chart") {
      col_var_chart = "value"
    } else if (base == "base_chart") {
      col_var_chart = "base_cat_label"
    } else if (base == "base_chart_unweighted") {
      col_var_chart = "base_cat_label_unweighted"
    }

  } else {
    if (base_mode == "Chart") {
      col_var_chart = "col_var_label"
    } else if (base == "base_chart") {
      col_var_chart = "col_var_label_with_base"
    } else if (base == "base_chart_unweighted") {
      col_var_chart = "col_var_label_with_base_unweighted"
    }
  }

  out[[col_var_chart]] <- factor(out[[col_var_chart]], levels = unique(out[[col_var_chart]])[order(1:length(unique(out[[col_var_chart]])), decreasing = label_direction)])
  out[["row_labels"]] <- factor(out[["row_labels"]], levels = unique(out[["row_labels"]])[order(1:length(unique(out[["row_labels"]])), decreasing = label_direction)])

  if (!sigtest) {
    chart <-  ms_barchart(data = out,
                          x = "row_labels",
                          y = "percentage",
                          group = col_var_chart)

    chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                        chart_direction = chart_direction,
                                        label_direction = label_direction) %>%
      chart_theme(legend_position = legend_pos)

  } else {
    chart <-  ms_barchart(data = out,
                          x = "row_labels",
                          y = "percentage",
                          group = col_var_chart,
                          labels = "sigPct")

    chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                        chart_direction = chart_direction,
                                        label_direction = label_direction) %>%
      chart_data_labels(show_val = FALSE, position = 'outEnd') %>%
      chart_theme(legend_position = legend_pos)
  }



  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(out[[base]]))),
                               layout = layout
  )

  pptx_obj

}


# ==============================================================================
# Add crosstab chart of grid variables to PPT
epinion_addslide_crosstab_grid = function(x, col_var, row_var, weight_var,
                                          slide_title, chart_title, slide_subtitle, footer,
                                          label_direction = FALSE, chart_direction,
                                          theme_color, base, legend_pos = "b",
                                          filter = "totalt=1", base_mode = "Chart", total_bar = "default",
                                          sort_by_cat = 'default', sort_order = TRUE, sigtest = FALSE,
                                          sig_level = 0.05, bonferroni = FALSE,
                                          pptx_obj, rp_analysis, layout) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  if (sort_by_cat != 'default' && !(sort_by_cat %in% levels(as_factor(mydata)[[row_var]]))) {
    stop(paste0("'", sort_by_cat, "' does not a value label of variable ", row_var, ". Please check!"))
  }


  # Prepare percentage dataframe
  df_chart <- epinion_calc_crosstab(x, col_var, row_var, weight_var,
                                   filter_var, filter_val,
                                   rp_analysis) %>%
    mutate(base_chart = round_half_up(expss::index_col(expss::match_col("#Total", row_labels), `Total|Count`), 0),
           base_chart_unweighted = round_half_up(expss::index_col(expss::match_col("#Total", row_labels), `Total|Count_unweight`), 0)
           ) %>%
    drop_na()

  df_count <- df_chart %>%
    as.data.frame(.) %>%
    select(., row_labels, ends_with("Count")) %>%
    rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
    filter(., row_labels == "#Total") %>%
    mutate(across(where(is.numeric), round_half_up)) %>%
    pivot_longer(., -row_labels, names_to = "value", values_to = "count")

  df_count_unweight <- df_chart %>%
    as.data.frame(.) %>%
    select(., row_labels, ends_with("Count_unweight")) %>%
    rename_with(~ sub("\\|Count_unweight$", "", .x), everything()) %>%
    filter(., row_labels == "#Total") %>%
    pivot_longer(., -row_labels, names_to = "value", values_to = "count_unweight")

  # sort option
  if (sort_by_cat == 'default' && is.na(sort_order)) {
    out <- df_chart %>%
      as.data.frame(.)
  } else if (sort_by_cat != 'default' & !is.na(sort_order)) {
    out <- df_chart %>%
      tab_transpose() %>%
      filter(., !grepl("\\|Count", row_labels)) %>%
      filter(., !grepl("base_chart", row_labels)) %>%
      mutate(group = ifelse(row_labels == "Total", 2, 1))

    if (sort_order == TRUE) {
      out <- out %>%
        group_by(group) %>%
        arrange(desc(!!as.symbol(sort_by_cat)), .by_group = TRUE)
    } else if (sort_order == FALSE) {
      out <- out %>%
        group_by(group) %>%
        arrange(!!as.symbol(sort_by_cat), .by_group = TRUE)
    }

    out <- out %>%
      ungroup() %>%
      drop_na() %>%
      select(., -group) %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      mutate(row_labels = row.names(.))
    names(out) <- out[1,]

    out <- out %>%
      filter(., row_labels != "row_labels") %>%
      mutate(across(!row_labels, as.numeric))

  }

  # Add total bar
  if (total_bar == "default") {
    out <- out %>%
      select(., row_labels, !contains("Count")) %>%
      select(., row_labels, !contains("base_"), -"Total")
  } else if (total_bar == "At first") {
    out <- out %>%
      select(., row_labels, !contains("Count")) %>%
      select(., row_labels, "Total", !contains("base_"))
  } else if (total_bar == "At last") {
    out <- out %>%
      select(., row_labels, !contains("Count")) %>%
      select(., row_labels, !contains("base_"))
  }

  # Calculate sigtest
  ctable_sigtest <- epinion_calc_crosstab(x, col_var, row_var, weight_var,
                                         filter_var, filter_val,
                                         rp_analysis,
                                         sigtest = TRUE,
                                         sig_level = sig_level,
                                         bonferroni = bonferroni)

  out <- out %>%
    filter(row_labels != "#Total") %>%
    mutate(across(where(is.numeric), ~./100)) %>%
    pivot_longer(., !row_labels, names_to = "value", values_to = "percentage") %>%
    left_join(., select(df_count, -row_labels), by = "value") %>%
    left_join(., select(df_count_unweight, -row_labels), by = "value") %>%
    mutate(base_cat_label = paste0(value, " (N=", trimws(round_half_up(count, 0)), ")"),
           base_cat_label_unweighted = paste0(value, " (N=", trimws(round_half_up(count_unweight, 0)), ")")
          ) %>%
    left_join(., select(as.data.frame(df_chart), row_labels, base_chart, base_chart_unweighted), by = "row_labels") %>%
    left_join(., ctable_sigtest, by = c("row_labels", "value")) %>%
    mutate(col_var_label = ifelse(value == "Total", value, paste0(value, " \n (", sigLetter, ")")),
           col_var_label_with_base = ifelse(value == "Total", base_cat_label, paste0(base_cat_label, " \n (", sigLetter, ")")),
           col_var_label_with_base_unweighted = ifelse(value == "Total", base_cat_label_unweighted, paste0(base_cat_label_unweighted, " \n (", sigLetter, ")")),
           sigPct = ifelse(value == "Total", paste0(as.character(round_half_up(percentage*100, 0)), "%"), sigPct)
          )

  if (!sigtest) {
    if (base_mode == "Chart") {
      col_var_chart = "value"
    } else if (base == "base_chart") {
      col_var_chart = "base_cat_label"
    } else if (base == "base_chart_unweighted") {
      col_var_chart = "base_cat_label_unweighted"
    }
  } else {
    if (base_mode == "Chart") {
      col_var_chart = "col_var_label"
    } else if (base == "base_chart") {
      col_var_chart = "col_var_label_with_base"
    } else if (base == "base_chart_unweighted") {
      col_var_chart = "col_var_label_with_base_unweighted"
    }
  }

  out[[col_var_chart]] <- factor(out[[col_var_chart]], levels = unique(out[[col_var_chart]])[order(1:length(unique(out[[col_var_chart]])), decreasing = label_direction)])
  out[["row_labels"]] <- factor(out[["row_labels"]], levels = levels(as_factor(x[[row_var]])))

  if (!sigtest) {
    chart <-  ms_barchart(data = out,
                          x = col_var_chart,
                          y = "percentage",
                          group = "row_labels")

    chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                        grouping = "stacked") %>%
      chart_theme(legend_position = legend_pos) %>%
      epinion_as_bar_stack(dir = chart_direction, gap_width = 100)
  } else {
    chart <-  ms_barchart(data = out,
                          x = col_var_chart,
                          y = "percentage",
                          group = "row_labels",
                          labels = "sigPct")

    chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                        grouping = "stacked") %>%
      chart_data_labels(show_val = FALSE) %>%
      chart_theme(legend_position = legend_pos) %>%
      epinion_as_bar_stack(dir = chart_direction, gap_width = 100)

  }

  # Prepare mean dataframe
  if (rp_analysis == "Crosstabs Row with Mean") {
    var_grid_mean <- paste0(row_var, "_mean")

    df_chart_mean <- epinion_calc_mean_crosstab(x, col_var, var_grid_mean, weight_var,
                                               filter_var, filter_val,
                                               rp_analysis) %>%
      pivot_longer(., -row_labels, names_to = "var", values_to = "statistic") %>%
      filter(grepl("Snitt", row_labels)) %>%
      mutate(Snitt = round_half_up(statistic, 2)) %>%
      mutate(row_labels = gsub("\\|Snitt", "", row_labels))

    if (sigtest) {
      df_chart_mean_sigtest <- epinion_calc_mean_crosstab(x, col_var, var_grid_mean, weight_var,
                                                          filter_var, filter_val,
                                                          rp_analysis,
                                                          sigtest = sigtest,
                                                          sig_level = sig_level,
                                                          bonferroni = bonferroni) %>%
        select(., -row_labels)

      df_chart_mean <- df_chart_mean %>%
        left_join(., df_chart_mean_sigtest, by = c("var" = "column_var")) %>%
        mutate(Snitt_sigtest = ifelse(var == "Total", as.character(round_half_up(Snitt, 2)), Snitt_sigtest))

    }

    if (total_bar == "default") {
      df_chart_mean <- df_chart_mean %>%
        filter(var != "Total")

    } else if (total_bar != "default") {
      # do nothing
    }


    # Sort option
    df_chart_mean <- df_chart_mean[match(unique(out$value), df_chart_mean$var),]

    # Create mean table
    if (!sigtest) {
      mean_table <- df_chart_mean[,c("Snitt")]
    } else {
      mean_table <- df_chart_mean %>%
        select(., Snitt_sigtest) %>%
        rename(Snitt = Snitt_sigtest)
    }

    ft <- epinion_format_table_PPT(mean_table) %>%
      width(width = 1.35) %>%
      height_all(height = (4.15-0.23-0.48)/(dim(mean_table)[1])) %>%
      height(height=0.23, part = "header")
    # 4.15 is the height of Content place holder in the slide template
    # 0.48 is the height for legend of the chart

  }

  if (rp_analysis == "Crosstabs Row with Mean") {
    pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                                 slide_title = slide_title,
                                 slide_subtitle = slide_subtitle,
                                 chart_title = chart_title,
                                 footer = footer,
                                 base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(out[[base]]))),
                                 layout = layout,
                                 mean_tbl_obj = ft)

  } else if (rp_analysis == "Crosstabs Row") {
    pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                                 slide_title = slide_title,
                                 slide_subtitle = slide_subtitle,
                                 chart_title = chart_title,
                                 footer = footer,
                                 base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(out[[base]]))),
                                 layout = layout)

  }


}


# ==============================================================================
# Add crosstab chart of multiple variables to PPT
epinion_addslide_crosstab_multiple = function(x, col_var, row_var, weight_var,
                                              slide_title, chart_title, slide_subtitle, footer,
                                              label_direction = FALSE, chart_direction,
                                              theme_color, base, legend_pos = "b",
                                              counted_value = "Yes", filter = "totalt=1", base_mode = "Chart", total_bar = "default",
                                              sort_by_cat = 'default', sort_order = TRUE, sigtest = FALSE,
                                              sig_level = 0.05, bonferroni = FALSE,
                                              pptx_obj, rp_analysis, layout) {

  var_multiple_row <- strsplit(row_var, "\\+")[[1]]
  var_multiple_col <- strsplit(col_var, "\\+")[[1]]


  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  if (length(var_multiple_col) >= 1 & length(var_multiple_row) > 1) {

    df_chart <- epinion_calc_crosstab(x, var_multiple_col, var_multiple_row, weight_var,
                                     filter_var, filter_val,
                                     rp_analysis,
                                     counted_value = counted_value) %>%
      as.data.frame(.) %>%
      mutate(base_chart = expss::index_col(expss::match_col("#Total", row_labels), `Total|Count`),
             base_chart_unweighted = expss::index_col(expss::match_col("#Total", row_labels), `Total|Count_unweight`))


    df_count <- df_chart %>%
      select(., row_labels, ends_with("Count")) %>%
      rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
      filter(., row_labels == "#Total") %>%
      mutate(across(where(is.numeric), round_half_up)) %>%
      pivot_longer(., -row_labels, names_to = "value", values_to = "count")

    df_count_unweight <- df_chart %>%
      select(., row_labels, ends_with("Count_unweight")) %>%
      rename_with(~ sub("\\|Count_unweight$", "", .x), everything()) %>%
      filter(., row_labels == "#Total") %>%
      pivot_longer(., -row_labels, names_to = "value", values_to = "count_unweight")

    # sort option
    if (sort_by_cat == 'default' & is.na(sort_order)) {
      df_chart <- df_chart %>%
        filter(row_labels != "#Total") %>%
        adorn_totals("row", fill = NA_integer_, na.rm = TRUE, name = "Total")

    } else if (sort_by_cat != 'default' & !is.na(sort_order)) {
      excluded_cats <- c(strsplit(sort_by_cat, "\\,")[[1]], "Total")
      df_sort_0 <- filter(df_chart, row_labels %in% excluded_cats)
      df_sort_1 <- filter(df_chart, !row_labels %in% excluded_cats)
      df_sort_1 <- df_sort_1[order(df_sort_1$Total, decreasing = sort_order), ]
      df_chart <- bind_rows(df_sort_1, df_sort_0) %>%
        filter(row_labels != "#Total") %>%
        adorn_totals("row", fill = NA_integer_, na.rm = TRUE, name = "Total")
    }

    out <- df_chart %>%
      select(., row_labels, !contains("Count")) %>%
      select(., row_labels, !contains("base_"), "Total") %>%
      filter(row_labels != "Total") %>%
      mutate(across(where(is.numeric), ~./100))

    # Add total bar
    if (total_bar == "default") {
      out <- out %>%
        select(., row_labels, !contains("Count")) %>%
        select(., row_labels, !contains("base_"), -"Total")
    } else if (total_bar == "At first") {
      out <- out %>%
        select(., row_labels, !contains("Count")) %>%
        select(., row_labels, "Total", !contains("base_"))
    } else if (total_bar == "At last") {
      # do nothing
    }

    out <- out %>%
      pivot_longer(., !row_labels, names_to = "value", values_to = "percentage") %>%
      left_join(., select(df_count, -row_labels), by = "value") %>%
      left_join(., select(df_count_unweight, -row_labels), by = "value") %>%
      mutate(base_cat_label = paste0(value, " (N=", trimws(round_half_up(count, 0)), ")"),
             base_cat_label_unweighted = paste0(value, " (N=", trimws(round_half_up(count_unweight, 0)), ")")
      ) %>%
      left_join(., select(df_chart, row_labels, base_chart, base_chart_unweighted), by = "row_labels")

    if (base_mode == "Chart") {
      col_var_chart = "value"
    } else if (base == "base_chart") {
      col_var_chart = "base_cat_label"
    } else if (base == "base_chart_unweighted") {
      col_var_chart = "base_cat_label_unweighted"
    }

    # Calculate sigtest
    if (sigtest) {
      ctable_sigtest <- epinion_calc_crosstab(x, var_multiple_col, var_multiple_row, weight_var,
                                              filter_var, filter_val,
                                              rp_analysis,
                                              counted_value = counted_value,
                                              sigtest = sigtest,
                                              sig_level = sig_level,
                                              bonferroni = bonferroni)

      out <- out %>%
        left_join(., ctable_sigtest, by = c("row_labels", "value")) %>%
        mutate(col_var_label = ifelse(value == "Total", value, paste0(value, " \n (", sigLetter, ")")),
               col_var_label_with_base = ifelse(value == "Total", base_cat_label, paste0(base_cat_label, " \n (", sigLetter, ")")),
               col_var_label_with_base_unweighted = ifelse(value == "Total", base_cat_label_unweighted, paste0(base_cat_label_unweighted, " \n (", sigLetter, ")")),
               sigPct = ifelse(value == "Total", paste0(as.character(round_half_up(percentage*100, 0)), "%"), sigPct)
              )

      if (base_mode == "Chart") {
        col_var_chart = "col_var_label"
      } else if (base == "base_chart") {
        col_var_chart = "col_var_label_with_base"
      } else if (base == "base_chart_unweighted") {
        col_var_chart = "col_var_label_with_base_unweighted"
      }

    }

    out[[col_var_chart]] <- factor(out[[col_var_chart]], levels = unique(out[[col_var_chart]])[order(1:length(unique(out[[col_var_chart]])), decreasing = label_direction)])
    out[["row_labels"]] <- factor(out[["row_labels"]], levels = unique(out[["row_labels"]])[order(1:length(unique(out[["row_labels"]])), decreasing = label_direction)])

    if (!sigtest) {
      chart <-  ms_barchart(data = out,
                            x = "row_labels",
                            y = "percentage",
                            group = col_var_chart)

      chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                          chart_direction = chart_direction,
                                          label_direction = label_direction) %>%
        chart_theme(legend_position = legend_pos)
    } else {
      chart <-  ms_barchart(data = out,
                            x = "row_labels",
                            y = "percentage",
                            group = col_var_chart,
                            labels = "sigPct")

      chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                          chart_direction = chart_direction,
                                          label_direction = label_direction) %>%
        chart_data_labels(show_val = FALSE, position = 'outEnd') %>%
        chart_theme(legend_position = legend_pos)
    }


  } else if (length(var_multiple_col) > 1 & length(var_multiple_row) == 1) {
    df_chart <- epinion_calc_crosstab(x, var_multiple_col, var_multiple_row, weight_var,
                                      filter_var, filter_val,
                                      rp_analysis,
                                      counted_value = counted_value) %>%
      mutate(base_chart = expss::index_col(expss::match_col("#Total", row_labels), `Total|Count`),
             base_chart_unweighted = expss::index_col(expss::match_col("#Total", row_labels), `Total|Count_unweight`)) %>%
      drop_na()


    df_count <- df_chart %>%
      as.data.frame(.) %>%
      select(., row_labels, ends_with("Count")) %>%
      rename_with(~ sub("\\|Count$", "", .x), everything()) %>%
      filter(., row_labels == "#Total") %>%
      mutate(across(where(is.numeric), round_half_up)) %>%
      pivot_longer(., -row_labels, names_to = "value", values_to = "count")

    df_count_unweight <- df_chart %>%
      as.data.frame(.) %>%
      select(., row_labels, ends_with("Count_unweight")) %>%
      rename_with(~ sub("\\|Count_unweight$", "", .x), everything()) %>%
      filter(., row_labels == "#Total") %>%
      pivot_longer(., -row_labels, names_to = "value", values_to = "count_unweight")

    # sort option
    if (sort_by_cat == 'default' && is.na(sort_order)) {
      out <- df_chart %>%
        as.data.frame(.)
    } else if (sort_by_cat != 'default' & !is.na(sort_order)) {
      out <- df_chart %>%
        tab_transpose() %>%
        as.data.frame(.) %>%
        filter(., !grepl("\\|Count", row_labels)) %>%
        filter(., !grepl("base_chart", row_labels)) %>%
        mutate(group = ifelse(row_labels == "Total", 2, 1))

      if (sort_order == TRUE) {
        out <- out %>%
        group_by(group) %>%
        arrange(desc(!!as.symbol(sort_by_cat)), .by_group = TRUE)
      } else if (sort_order == FALSE) {
        out <- out %>%
          group_by(group) %>%
          arrange(!!as.symbol(sort_by_cat), .by_group = TRUE)
      }

      out <- out %>%
        ungroup() %>%
        select(., -group) %>%
        tab_transpose() %>%
        as.data.frame(.) %>%
        mutate(row_labels = row.names(.))
      names(out) <- out[1,]

      out <- out %>%
        filter(., row_labels != "row_labels") %>%
        mutate(across(!row_labels, as.numeric))
    }

    # Add total bar
    if (total_bar == "default") {
      out <- out %>%
        select(., row_labels, !contains("Count")) %>%
        select(., row_labels, !contains("base_"), -"Total")
    } else if (total_bar == "At first") {
      out <- out %>%
        select(., row_labels, !contains("Count")) %>%
        select(., row_labels, "Total", !contains("base_"))
    } else if (total_bar == "At last") {
      out <- out %>%
        select(., row_labels, !contains("Count")) %>%
        select(., row_labels, !contains("base_"))
    }

    out <- out %>%
      filter(row_labels != "#Total") %>%
      mutate(across(where(is.numeric), ~./100)) %>%
      pivot_longer(., !row_labels, names_to = "value", values_to = "percentage") %>%
      left_join(., select(df_count, -row_labels), by = "value") %>%
      left_join(., select(df_count_unweight, -row_labels), by = "value") %>%
      mutate(base_cat_label = paste0(value, " (N=", trimws(round_half_up(count, 0)), ")"),
             base_cat_label_unweighted = paste0(value, " (N=", trimws(round_half_up(count_unweight, 0)), ")")
             ) %>%
      left_join(., select(as.data.frame(df_chart), row_labels, base_chart, base_chart_unweighted), by = "row_labels")

    if (base_mode == "Chart") {
      col_var_chart = "value"
    } else if (base == "base_chart") {
      col_var_chart = "base_cat_label"
    } else if (base == "base_chart_unweighted") {
      col_var_chart = "base_cat_label_unweighted"
    }

    if (sigtest) {
      # Calculate sigtest
      ctable_sigtest <- epinion_calc_crosstab(x, var_multiple_col, var_multiple_row, weight_var,
                                              filter_var, filter_val,
                                              rp_analysis,
                                              counted_value = counted_value,
                                              sigtest = sigtest,
                                              sig_level = sig_level,
                                              bonferroni = bonferroni)

      out <- out %>%
        left_join(., ctable_sigtest, by = c("row_labels", "value")) %>%
        mutate(col_var_label = ifelse(value == "Total", value, paste0(value, " \n (", sigLetter, ")")),
               col_var_label_with_base = ifelse(value == "Total", base_cat_label, paste0(base_cat_label, " \n (", sigLetter, ")")),
               col_var_label_with_base_unweighted = ifelse(value == "Total", base_cat_label_unweighted, paste0(base_cat_label_unweighted, " \n (", sigLetter, ")")),
               sigPct = ifelse(value == "Total", paste0(as.character(round_half_up(percentage*100, 0)), "%"), sigPct)
              )

      if (base_mode == "Chart") {
        col_var_chart = "col_var_label"
      } else if (base == "base_chart") {
        col_var_chart = "col_var_label_with_base"
      } else if (base == "base_chart_unweighted") {
        col_var_chart = "col_var_label_with_base_unweighted"
      }

    }

    if (rp_analysis == "Multiple Crosstabs (By Case)") {
      out[[col_var_chart]] <- factor(out[[col_var_chart]], levels = unique(out[[col_var_chart]])[order(1:length(unique(out[[col_var_chart]])), decreasing = label_direction)])
      out[["row_labels"]] <- factor(out[["row_labels"]], levels = unique(out[["row_labels"]])[order(1:length(unique(out[["row_labels"]])), decreasing = label_direction)])

      if (!sigtest) {
        chart <-  ms_barchart(data = out,
                              x = "row_labels",
                              y = "percentage",
                              group = col_var_chart)

        chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                            chart_direction = chart_direction,
                                            label_direction = label_direction) %>%
          chart_theme(legend_position = legend_pos)
      } else {
        chart <-  ms_barchart(data = out,
                              x = "row_labels",
                              y = "percentage",
                              group = col_var_chart,
                              labels = "sigPct")

        chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                            chart_direction = chart_direction,
                                            label_direction = label_direction) %>%
          chart_data_labels(show_val = FALSE, position = 'outEnd') %>%
          chart_theme(legend_position = legend_pos)
      }

    } else if (grepl("Multiple Crosstabs Row (By Case)", rp_analysis, fixed = TRUE)) {
      out[[col_var_chart]] <- factor(out[[col_var_chart]], levels = unique(out[[col_var_chart]])[order(1:length(unique(out[[col_var_chart]])), decreasing = label_direction)])
      out[["row_labels"]] <- factor(out[["row_labels"]], levels = levels(as_factor(x[[var_multiple_row]])))

      if (!sigtest) {
        chart <-  ms_barchart(data = out,
                              x = col_var_chart,
                              y = "percentage",
                              group = "row_labels")
        chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                            label_direction = label_direction,
                                            grouping = "stacked") %>%
          chart_theme(legend_position = legend_pos) %>%
          epinion_as_bar_stack(dir = chart_direction, gap_width = 100)

      } else {
        chart <-  ms_barchart(data = out,
                              x = col_var_chart,
                              y = "percentage",
                              group = "row_labels",
                              labels = "sigPct")
        chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                            label_direction = label_direction,
                                            grouping = "stacked") %>%
          chart_theme(legend_position = legend_pos) %>%
          chart_data_labels(show_val = FALSE) %>%
          epinion_as_bar_stack(dir = chart_direction, gap_width = 100)

      }


      # Prepare mean dataframe
      if (rp_analysis == "Multiple Crosstabs Row (By Case) with Mean") {
        var_grid_mean <- paste0(var_multiple_row, "_mean")

        df_chart_mean <- epinion_calc_mean_crosstab(x, var_multiple_col, var_grid_mean, weight_var,
                                                   filter_var, filter_val,
                                                   rp_analysis,
                                                   counted_value = counted_value) %>%
          pivot_longer(., -row_labels, names_to = "var", values_to = "statistic") %>%
          filter(grepl("Snitt", row_labels)) %>%
          mutate(Snitt = round_half_up(statistic, 2)) %>%
          mutate(row_labels = gsub("\\|Snitt", "", row_labels))

        if (sigtest) {
          df_chart_mean_sigtest <- epinion_calc_mean_crosstab(x, var_multiple_col, var_grid_mean, weight_var,
                                                              filter_var, filter_val,
                                                              rp_analysis,
                                                              counted_value = counted_value,
                                                              sigtest = sigtest,
                                                              sig_level = sig_level,
                                                              bonferroni = bonferroni) %>%
            select(., -row_labels)

          df_chart_mean <- df_chart_mean %>%
            left_join(., df_chart_mean_sigtest, by = c("var" = "column_var")) %>%
            mutate(Snitt_sigtest = ifelse(var == "Total", as.character(round_half_up(Snitt, 2)), Snitt_sigtest))

        }

        if (total_bar == "default") {
          df_chart_mean <- df_chart_mean %>%
            filter(var != "Total")

        } else if (total_bar != "default") {
          # do nothing
        }

        # Sort option
        df_chart_mean <- df_chart_mean[match(unique(out$value), df_chart_mean$var),]

        # Create mean table
        if (!sigtest) {
          mean_table <- df_chart_mean[,c("Snitt")]

        } else {
          mean_table <- df_chart_mean %>%
            select(., Snitt_sigtest) %>%
            rename(Snitt = Snitt_sigtest)

        }

        ft <- epinion_format_table_PPT(mean_table) %>%
          width(width = 1.35) %>%
          height_all(height = (4.15-0.23-0.48)/(dim(mean_table)[1])) %>%
          height(height=0.23, part = "header")
        # 4.15 is the height of Content place holder in the slide template
        # 0.48 is the height for legend of the chart

      }

    }

  }


  if (rp_analysis == "Multiple Crosstabs Row (By Case) with Mean") {
    pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                                 slide_title = slide_title,
                                 slide_subtitle = slide_subtitle,
                                 chart_title = chart_title,
                                 footer = footer,
                                 base_chart = ifelse(base_mode == "Category", "", paste0("N=", min(unique(round_half_up(out[[base]]),0), na.rm = TRUE))),
                                 layout = layout,
                                 mean_tbl_obj = ft)
  } else {
    pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                                 slide_title = slide_title,
                                 slide_subtitle = slide_subtitle,
                                 chart_title = chart_title,
                                 footer = footer,
                                 base_chart = ifelse(base_mode == "Category", "", paste0("N=", min(unique(round_half_up(out[[base]]),0), na.rm = TRUE))),
                                 layout = layout)
  }

}


# ==============================================================================
# Add crosstab mean bar chart of single variables to PPT
epinion_addslide_meanchart_crosstab = function(x, col_var, row_var, weight_var,
                                                    slide_title, chart_title, slide_subtitle, footer,
                                                    label_direction = FALSE, chart_direction,
                                                    theme_color, base,
                                                    counted_value = "Yes", filter = "totalt=1", base_mode = "Chart", total_bar = "default",
                                                    sort_by_cat = 'default', sort_order = TRUE, sigtest = FALSE,
                                                    sig_level = 0.05, bonferroni = FALSE,
                                                    pptx_obj, rp_analysis, layout) {

  var_multiple_row <- strsplit(row_var, "\\+")[[1]]
  var_multiple_col <- strsplit(col_var, "\\+")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  df_chart <- epinion_calc_mean_crosstab(x, var_multiple_col, row_var, weight_var,
                                         filter_var, filter_val,
                                         rp_analysis,
                                         counted_value = counted_value) %>%
    pivot_longer(., -row_labels, names_to = "var", values_to = "statistic") %>%
    pivot_wider(., names_from = "row_labels", values_from = "statistic") %>%
    mutate(base_chart = round_half_up(max(`Valid N`, na.rm = TRUE),0),
           base_chart_unweighted = round_half_up(max(`Unw. valid N`, na.rm = TRUE),0),
           base_cat_label = paste0(var, " (N=", trimws(round_half_up(`Valid N`, 0)), ")"),
           base_cat_label_unweighted = paste0(var, " (N=", trimws(round_half_up(`Unw. valid N`, 0)), ")"),
           totalt = 1) %>%
    rename(row_labels = var,
           Mean = Snitt)

  if (base_mode == "Chart") {
    row_var_chart = "row_labels"
  } else if (base == "base_chart") {
    row_var_chart = "base_cat_label"
  } else if (base == "base_chart_unweighted") {
    row_var_chart = "base_cat_label_unweighted"
  }

  # Calculate sigtest
  if (sigtest) {

    ctable_sigtest <- epinion_calc_mean_crosstab(x, var_multiple_col, var_multiple_row, weight_var,
                                                filter_var, filter_val,
                                                rp_analysis,
                                                counted_value = counted_value,
                                                sigtest = sigtest,
                                                sig_level = sig_level,
                                                bonferroni = bonferroni)

    df_chart <- df_chart %>%
      left_join(., ctable_sigtest, by = c("row_labels" = "column_var")) %>%
      mutate(Snitt = ifelse(row_labels == "Total", as.character(round_half_up(Mean, 2)), Snitt_sigtest),
             col_var_label = ifelse(row_labels == "Total", row_labels, paste0(row_labels, " \n (", sigLetter, ")")),
             col_var_label_with_base = ifelse(row_labels == "Total", base_cat_label, paste0(base_cat_label, " \n (", sigLetter, ")")),
             col_var_label_with_base_unweighted = ifelse(row_labels == "Total", base_cat_label_unweighted, paste0(base_cat_label_unweighted, " \n (", sigLetter, ")"))
             )

    if (base_mode == "Chart") {
      row_var_chart = "col_var_label"
    } else if (base == "base_chart") {
      row_var_chart = "col_var_label_with_base"
    } else if (base == "base_chart_unweighted") {
      row_var_chart = "col_var_label_with_base_unweighted"
    }

  }

  # sort option
  if (sort_by_cat == 'default' & is.na(sort_order)) {
    # do nothing

  } else if (!is.na(sort_order)) {
    excluded_cats <- c(strsplit(sort_by_cat, "\\,")[[1]], "Total")
    df_sort_0 <- filter(df_chart, row_labels %in% excluded_cats)
    df_sort_1 <- filter(df_chart, !row_labels %in% excluded_cats)
    df_sort_1 <- df_sort_1[order(df_sort_1$Mean, decreasing = sort_order), ]
    df_chart <- bind_rows(df_sort_1, df_sort_0)

  }

  # Add total bar
  if (total_bar == "default") {
    df_chart <- df_chart %>%
      filter(row_labels != "Total")

  } else if (total_bar == "At first") {
    df_chart <- df_chart[c(nrow(df_chart),1:nrow(df_chart)-1),]

  } else if (total_bar == "At last") {
    # do nothing
  }

  df_chart[[row_var_chart]] <- factor(df_chart[[row_var_chart]], levels = df_chart[[row_var_chart]][order(1:length(df_chart[[row_var_chart]]), decreasing = label_direction)])

  if (!sigtest) {
    chart <-  ms_barchart(data = df_chart,
                          x = row_var_chart,
                          y = "Mean",
                          group = "totalt")

    chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                        chart_direction = chart_direction,
                                        label_direction = label_direction) %>%
      chart_theme(legend_position = "n") %>%
      chart_ax_y(display = 0,
                 limit_min = 0,
                 limit_max = ifelse(max(df_chart$Mean) < 5, 5, max(df_chart$Mean) + 1),
                 num_fmt = "0.00") %>%
      chart_data_labels(show_val = TRUE,
                        position = "outEnd",
                        num_fmt = "0.00")

  } else {
    chart <-  ms_barchart(data = df_chart,
                          x = row_var_chart,
                          y = "Mean",
                          group = "totalt",
                          labels = "Snitt")

    chart <- epinion_barchart_formating(chart, colour_list = theme_color,
                                        chart_direction = chart_direction,
                                        label_direction = label_direction) %>%
      chart_theme(legend_position = "n") %>%
      chart_ax_y(display = 0,
                 limit_min = 0,
                 limit_max = ifelse(max(df_chart$Mean) < 5, 5, max(df_chart$Mean) + 1),
                 num_fmt = "0.00") %>%
      chart_data_labels(show_val = FALSE,
                        position = "outEnd")

  }


  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = ifelse(base_mode == "Category", "", paste0("N=", unique(max(df_chart[[base]])))),
                               layout = layout)

}


# ##############################################################################
# Line chart
# ##############################################################################
# Formating for line chart
epinion_linechart_formating = function (chart_obj, colour_list,
                                        number_format = "General") {

  chart_obj <- chart_obj %>%
    chart_theme(legend_text = fp_text(color = "#0f283c", font.size = 10),
                axis_text_x = fp_text(color = "#0f283c", font.size = 10),
                axis_text_y = fp_text(color = "#0f283c", font.size = 10),
                grid_major_line_x = fp_border(width = 0),
                grid_major_line_y = fp_border(width = 0)) %>%
    chart_ax_x(num_fmt = number_format,
               major_tick_mark = 'none') %>%
    chart_ax_y(display = 0,
               limit_min = 1,
               limit_max = 5,
               num_fmt = "General") %>%
    # enable chart data labels
    chart_data_labels(show_val = TRUE,
                      position = "t",
                      num_fmt = "0.00") %>%
    # hide chart title
    chart_labels()

  series_names <- mschart:::get_series_names(chart_obj)

  if (length(colour_list) > 1) {
    if(length(series_names) <= length(colour_list)) {
      palette_ <- colour_list[[length(series_names)]]
    } else {
      palette_ <- sample(colors(), size = length(series_names), replace = TRUE)
    }
  } else {
    palette_ <- colour_list[[1]]
  }

  series_symbols <- rep("none", length(series_names))
  series_lstyle <- rep("solid", length(series_names))
  series_lwidth <- rep(2, length(series_names))
  labels_fp <- rep(list(fp_text(color = "#0f283c", font.size = 10)), length(series_names))
  series_smooth <- rep(0,length(series_names))

  chart_obj$series_settings <- list(
    fill = setNames(palette_, series_names),
    colour = setNames(palette_, series_names),
    symbol = setNames(series_symbols, series_names),
    line_style = setNames(series_lstyle, series_names),
    line_width = setNames(series_lwidth, series_names),
    labels_fp = setNames(labels_fp, series_names),
    smooth = setNames(series_smooth, series_names)
  )

  chart_obj
}

# ==============================================================================
# Add frequency/crosstab line chart of single variable to PPT
epinion_addslide_linechart_cro = function(x, col_var, row_var, weight_var, group_var,
                                          slide_title, chart_title, slide_subtitle, footer,
                                          theme_color, base, legend_pos = "b",
                                          counted_value = "Yes",
                                          label_direction = FALSE,
                                          filter = "totalt=1", total_line = "default",
                                          sigtest = FALSE,
                                          sig_level = 0.05, bonferroni = FALSE,
                                          pptx_obj, rp_analysis, layout) {

  var_multiple_col <- strsplit(col_var, "\\+")[[1]]
  var_multiple_group <- strsplit(group_var, "\\+")[[1]]

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  if (length(var_multiple_col) == 1) {
    if (length(var_multiple_group) == 1) {
      df_chart_input <- x %>%
        mutate(totalt = 1) %>%
        mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
        filter(filter_tmp == filter_val) %>%
        filter(!is.na(!!as.symbol(col_var))) %>%
        tab_cells("|" = unvr(..$row_var ))  %>%
        tab_cols("|" = unvr(..$col_var ) %nest% ..$group_var,
                 total(label = "Total") %nest% ..$group_var)

    } else if (length(var_multiple_group) > 1) {
      names_old <- unique(unlist(as_factor(x)[, var_multiple_group]))
      names_new <- ifelse(names_old == counted_value, 1, 0)
      df_conversion <- data.frame(names_old, names_new)

      var_multiple_group_tmp <- x %>%
        mutate(totalt = 1) %>%
        mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
        filter(filter_tmp == filter_val) %>%
        as_factor(., only_labelled = TRUE) %>%
        mutate(id_tmp = row.names(.)) %>%
        select(., id_tmp, all_of(var_multiple_group)) %>%
        mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

      df_chart_input <- x %>%
        mutate(totalt = 1) %>%
        mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
        filter(filter_tmp == filter_val) %>%
        select(., !all_of(var_multiple_group)) %>%
        mutate(id_tmp = row.names(.)) %>%
        left_join(., var_multiple_group_tmp, by = "id_tmp") %>%
        select(., -id_tmp) %>%
        set_variable_labels(., .labels = var_label(x)) %>%
        filter(!is.na(!!as.symbol(col_var))) %>%
        tab_cells("|" = unvr(..$row_var ))  %>%
        tab_cols("|" = unvr(..$col_var ) %nest% mdset(..$head(var_multiple_group, 1) %to% ..$tail(var_multiple_group, 1)),
                 total(label = "Total") %nest% mdset(..$head(var_multiple_group, 1) %to% ..$tail(var_multiple_group, 1)),
                 totalt)

    }

    df_chart <- df_chart_input %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_mean_sd_n(weighted_valid_n = TRUE) %>%
      tab_weight(weight = total()) %>%
      tab_stat_mean_sd_n(weighted_valid_n = FALSE,
                         labels = c("Mean_unweight", "Std. dev. unweight", "Unw. valid N")) %>%
      tab_pivot() %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      separate(row_labels, c("column_var", "group"), sep = "\\|", remove = TRUE) %>%
      mutate(base_chart = ifelse(length(var_multiple_group) == 1,
                                 sum(`Valid N`, na.rm = TRUE)/2,
                                 expss::index_col(expss::match_col("totalt", column_var), `Valid N`)),
             base_chart = round_half_up(base_chart, 0),
             base_chart_unweighted = ifelse(length(var_multiple_group) == 1,
                                            round_half_up(sum(`Unw. valid N`, na.rm = TRUE),0)/2,
                                            expss::index_col(expss::match_col("totalt", column_var), `Unw. valid N`)
                                            )
            ) %>%
      filter(column_var != "totalt")

    if (sigtest) {
      if (group_var != "totalt") {
        ctable_sigtest <- df_chart_input %>%
          tab_weight(weight = ..$weight_var) %>%
          tab_stat_mean_sd_n(labels = c("Mean_sigtest", "Std. dev.", "Unw. valid N")) %>%
          tab_last_sig_means(sig_level = sig_level, bonferroni = bonferroni,
                             digits = 2, subtable_marks = "greater") %>%
          tab_pivot() %>%
          tab_transpose() %>%
          as.data.frame(.) %>%
          separate(row_labels, c("column_var", "group", "sigLetter"), sep = "\\|", remove = TRUE) %>%
          select(., -`Std. dev.`, -`Unw. valid N`)

        df_chart <- df_chart %>%
          left_join(., ctable_sigtest, by = c("column_var", "group")) %>%
          mutate(group_sigtest = paste0(group, " (", sigLetter, ")"))

      } else if (group_var == "totalt") {
        ctable_sigtest <- epinion_calc_mean_crosstab(x, col_var, row_var, weight_var,
                                                    filter_var, filter_val,
                                                    rp_analysis,
                                                    sigtest = sigtest,
                                                    sig_level = sig_level,
                                                    bonferroni = bonferroni)

        df_chart <- df_chart %>%
          left_join(., ctable_sigtest, by = c("column_var")) %>%
          mutate(group_sigtest = group,
                 Mean_sigtest = ifelse(column_var == "Total", as.character(round_half_up(Mean, 2)), Snitt_sigtest)) %>%
          rename(column_var_org = column_var) %>%
          mutate(column_var = ifelse(column_var_org == "Total", "Total", paste0(column_var_org, " \n (", sigLetter, ")")))

      }

    }

  } else if (length(var_multiple_col) > 1) {
    names_old <- unique(unlist(as_factor(x)[, var_multiple_col]))
    names_new <- ifelse(names_old == counted_value, 1, 0)
    df_conversion <- data.frame(names_old, names_new)

    var_multiple_col_tmp <- x %>%
      mutate(totalt = 1) %>%
      mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
      filter(filter_tmp == filter_val) %>%
      as_factor(., only_labelled = TRUE) %>%
      mutate(id_tmp = row.names(.)) %>%
      select(., id_tmp, all_of(var_multiple_col)) %>%
      mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

    if (length(var_multiple_group) == 1) {
      df_chart_input <- x %>%
        mutate(totalt = 1) %>%
        mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
        filter(filter_tmp == filter_val) %>%
        select(., !all_of(var_multiple_col)) %>%
        mutate(id_tmp = row.names(.)) %>%
        left_join(., var_multiple_col_tmp, by = "id_tmp") %>%
        select(., -id_tmp) %>%
        set_variable_labels(., .labels = var_label(x)) %>%
        select(., totalt, !!as.symbol(weight_var), !!as.symbol(row_var), !!as.symbol(group_var), all_of(var_multiple_col)) %>%
        filter(if_any(var_multiple_col[1]:var_multiple_col[length(var_multiple_col)], ~. == 1)) %>%
        tab_cells("|" = unvr(..$row_var ))  %>%
        tab_cols(mdset(..$var_multiple_col[1] %to% ..$var_multiple_col[length(var_multiple_col)]) %nest% ..$group_var,
                 total(label = "Total") %nest% ..$group_var)

    } else if (length(var_multiple_group) > 1) {
      var_multiple_group_tmp <- x %>%
        mutate(totalt = 1) %>%
        mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
        filter(filter_tmp == filter_val) %>%
        as_factor(., only_labelled = TRUE) %>%
        mutate(id_tmp = row.names(.)) %>%
        select(., id_tmp, all_of(var_multiple_group)) %>%
        mutate(across(!id_tmp, ~ deframe(df_conversion)[.]))

      df_chart_input <- x %>%
        mutate(totalt = 1) %>%
        mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
        filter(filter_tmp == filter_val) %>%
        select(., !all_of(var_multiple_col)) %>%
        select(., !all_of(var_multiple_group)) %>%
        mutate(id_tmp = row.names(.)) %>%
        left_join(., var_multiple_col_tmp, by = "id_tmp") %>%
        left_join(., var_multiple_group_tmp, by = "id_tmp") %>%
        select(., -id_tmp) %>%
        set_variable_labels(., .labels = var_label(x)) %>%
        select(., totalt, !!as.symbol(weight_var), !!as.symbol(row_var), all_of(var_multiple_group), all_of(var_multiple_col)) %>%
        filter(if_any(var_multiple_col[1]:var_multiple_col[length(var_multiple_col)], ~. == 1)) %>%
        tab_cells("|" = unvr(..$row_var ))  %>%
        tab_cols(mdset(..$var_multiple_col[1] %to% ..$var_multiple_col[length(var_multiple_col)]) %nest% mdset(..$head(var_multiple_group, 1) %to% ..$tail(var_multiple_group, 1)),
                 total(label = "Total") %nest% mdset(..$head(var_multiple_group, 1) %to% ..$tail(var_multiple_group, 1)))

    }

    df_chart <- df_chart_input %>%
      tab_weight(weight = ..$weight_var) %>%
      tab_stat_mean_sd_n(weighted_valid_n = TRUE) %>%
      tab_weight(weight = total()) %>%
      tab_stat_mean_sd_n(weighted_valid_n = FALSE,
                         labels = c("Mean_unweight", "Std. dev. unweight", "Unw. valid N")) %>%
      tab_pivot() %>%
      tab_transpose() %>%
      as.data.frame(.) %>%
      separate(row_labels, c("column_var", "group"), sep = "\\|", remove = TRUE) %>%
      mutate(tmp = ifelse(column_var == "Total", 2, 1)) %>%
      group_by(tmp) %>%
      mutate(base_chart = round_half_up(sum(`Valid N`, na.rm = TRUE),0),
             base_chart_unweighted = round_half_up(sum(`Unw. valid N`, na.rm = TRUE),0)) %>%
      ungroup() %>%
      mutate(base_chart = expss::index_col(expss::match_col("Total", column_var), base_chart),
             base_chart_unweighted = expss::index_col(expss::match_col("Total", column_var), base_chart_unweighted))

    if (sigtest) {
      if (group_var != "totalt") {
        ctable_sigtest <- df_chart_input %>%
          tab_weight(weight = ..$weight_var) %>%
          tab_stat_mean_sd_n(labels = c("Mean_sigtest", "Std. dev.", "Unw. valid N")) %>%
          tab_last_sig_means(sig_level = sig_level, bonferroni = bonferroni,
                             digits = 2, subtable_marks = "greater") %>%
          tab_pivot() %>%
          tab_transpose() %>%
          as.data.frame(.) %>%
          separate(row_labels, c("column_var", "group", "sigLetter"), sep = "\\|", remove = TRUE) %>%
          select(., -`Std. dev.`, -`Unw. valid N`)

        df_chart <- df_chart %>%
          left_join(., ctable_sigtest, by = c("column_var", "group")) %>%
          mutate(group_sigtest = paste0(group, " (", sigLetter, ")"))

      } else if (group_var == "totalt") {
        ctable_sigtest <- epinion_calc_mean_crosstab(x, var_multiple_col, row_var, weight_var,
                                                     filter_var, filter_val,
                                                     rp_analysis,
                                                     sigtest = sigtest,
                                                     sig_level = sig_level,
                                                     bonferroni = bonferroni)

        df_chart <- df_chart %>%
          left_join(., ctable_sigtest, by = c("column_var")) %>%
          mutate(group_sigtest = group,
                 Mean_sigtest = ifelse(column_var == "Total", as.character(round_half_up(Mean, 2)), Snitt_sigtest)) %>%
          rename(column_var_org = column_var) %>%
          mutate(column_var = ifelse(column_var_org == "Total", "Total", paste0(column_var_org, " \n (", sigLetter, ")")))

      }

    }

  }

  # Add total line
  if (total_line == "default") {
    df_chart <- df_chart %>%
      filter(column_var != "Total")

  } else if (total_line == "At first") {
    suppressMessages(if (!require("collapse", quietly = TRUE)) install.packages("collapse", quiet = TRUE))
    df_chart <- collapse::roworderv(df_chart, neworder = df_chart$column_var == "Total")

  } else if (total_line == "At last") {
    # do nothing
  }

  df_chart[["column_var"]] <- factor(df_chart[["column_var"]], levels = unique(df_chart[["column_var"]])[order(1:length(unique(df_chart[["column_var"]])), decreasing = label_direction)])

  if (!sigtest) {
    if (group_var != "totalt") {
      df_chart[["group"]] <- factor(df_chart[["group"]], levels = unique(df_chart[["group"]])[order(1:length(unique(df_chart[["group"]])), decreasing = label_direction)])

    }


    chart <-  ms_linechart(data = df_chart,
                           x = "column_var",
                           y = "Mean",
                           group = "group")

    chart <- epinion_linechart_formating(chart, colour_list = theme_color) %>%
      chart_theme(legend_position = legend_pos)

  } else {
    if (group_var != "totalt") {
      df_chart[["group_sigtest"]] <- factor(df_chart[["group_sigtest"]], levels = unique(df_chart[["group_sigtest"]])[order(1:length(unique(df_chart[["group_sigtest"]])), decreasing = label_direction)])

    }

    chart <-  ms_linechart(data = df_chart,
                           x = "column_var",
                           y = "Mean",
                           group = "group_sigtest",
                           labels = "Mean_sigtest")

    chart <- epinion_linechart_formating(chart, colour_list = theme_color) %>%
      chart_theme(legend_position = legend_pos) %>%
      chart_data_labels(show_val = FALSE,
                        position = "t")

  }


  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = paste0("N=", unique(df_chart[[base]])),
                               layout = layout)

}

# ==============================================================================
# Add line chart of grid variables to PPT
epinion_addslide_linechart_cro_grid = function(x, col_var, row_var, weight_var,
                                               slide_title, chart_title, slide_subtitle, footer,
                                               theme_color, base, legend_pos = "b",
                                               filter = "totalt=1",
                                               counted_value = "Yes",
                                               total_line = "default",
                                               sigtest = FALSE,
                                               sig_level = 0.05, bonferroni = FALSE,
                                               pptx_obj, rp_analysis, layout) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  var_grid <- strsplit(row_var, "\\,")[[1]]
  var_multiple_col <- strsplit(col_var, "\\+")[[1]]

  df_chart <- epinion_calc_mean_crosstab(x, var_multiple_col, var_grid, weight_var,
                                         filter_var, filter_val,
                                         rp_analysis,
                                         counted_value = counted_value)

  # Add total line
  if (total_line == "At first") {
    df_chart <- df_chart %>%
      select(., row_labels, Total, colnames(.))

  } else if (total_line == "At last") {
    # do nothing
  }

  df_chart <- df_chart %>%
    separate(., row_labels, c("var_labels", "statistic"), sep = "\\|", remove = TRUE) %>%
    filter(!grepl("Std. dev.", statistic)) %>%
    pivot_longer(., -c(var_labels, statistic), names_to = "column_var") %>%
    pivot_wider(., names_from = statistic, values_from = c(value)) %>%
    mutate(tmp = ifelse(column_var == "Total", 2, 1)) %>%
    group_by(tmp) %>%
    mutate(base_chart = round_half_up(max(`Valid N`, na.rm = TRUE), 0),
           base_chart_unweighted = round_half_up(max(`Unw. valid N`, na.rm = TRUE),0)) %>%
    ungroup() %>%
    mutate(base_chart = expss::index_col(expss::match_col("Total", column_var), base_chart),
           base_chart_unweighted = expss::index_col(expss::match_col("Total", column_var), base_chart_unweighted))

  if (total_line == "default") {
    df_chart <- df_chart %>%
      filter(column_var != "Total")
  }

  if (sigtest) {
    ctable_sigtest <- epinion_calc_mean_crosstab(x, var_multiple_col, var_grid, weight_var,
                                                filter_var, filter_val,
                                                rp_analysis,
                                                counted_value = counted_value,
                                                sigtest = sigtest,
                                                sig_level = sig_level,
                                                bonferroni = bonferroni)

    df_chart <- df_chart %>%
      left_join(., ctable_sigtest, by = c("var_labels" = "row_labels", "column_var")) %>%
      rename(column_var_org = column_var) %>%
      mutate(Mean_sigtest = ifelse(column_var_org == "Total", as.character(round_half_up(Mean, 2)), Snitt_sigtest),
             column_var = ifelse(column_var_org == "Total", "Total", paste0(column_var_org, " \n (", sigLetter, ")"))
            )

  }

  df_chart[["var_labels"]] <- factor(df_chart[["var_labels"]], levels = unique(df_chart[["var_labels"]])[order(1:length(unique(df_chart[["var_labels"]])), decreasing = FALSE)])
  df_chart[["column_var"]] <- factor(df_chart[["column_var"]], levels = unique(df_chart[["column_var"]])[order(1:length(unique(df_chart[["column_var"]])), decreasing = FALSE)])

  if (!sigtest) {
    chart <-  ms_linechart(data = df_chart,
                           x = "column_var",
                           y = "Mean",
                           group = "var_labels")

    chart <- epinion_linechart_formating(chart, colour_list = theme_color) %>%
      chart_theme(legend_position = legend_pos)

  } else {
    chart <-  ms_linechart(data = df_chart,
                           x = "column_var",
                           y = "Mean",
                           group = "var_labels",
                           labels = "Mean_sigtest")

    chart <- epinion_linechart_formating(chart, colour_list = theme_color) %>%
      chart_theme(legend_position = legend_pos) %>%
      chart_data_labels(show_val = FALSE,
                        position = "t")

  }


  pptx_obj <- epinion_addslide(pptx_obj, chart, rp_analysis,
                               slide_title = slide_title,
                               slide_subtitle = slide_subtitle,
                               chart_title = chart_title,
                               footer = footer,
                               base_chart = paste0("N=", unique(df_chart[[base]])),
                               layout = layout)

}

# ##############################################################################
# Open answer
# ##############################################################################
# Add open answer slide to PPT
epinion_addslide_openans = function(x, open_ans_var,
                                    slide_title, chart_title, slide_subtitle, footer,
                                    filter = "totalt=1", report_mode,
                                    pptx_obj, rp_analysis, layout) {

  filter_var <- do.call(rbind, strsplit(filter, split = "="))[1]
  filter_val <- do.call(rbind, strsplit(filter, split = "="))[2]

  if (report_mode == "unique") {
    open_answer <- x %>%
      mutate(totalt = 1) %>%
      mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
      filter(filter_tmp == filter_val) %>%
      arrange(!!as.symbol(open_ans_var)) %>%
      select(., !!as.symbol(open_ans_var)) %>%
      filter(!!as.symbol(open_ans_var) != "") %>%
      group_by(!!as.symbol(open_ans_var)) %>%
      summarise(count = n()) %>%
      mutate(base_chart = sum(count, na.rm = TRUE))

  } else if (report_mode == "full") {
    open_answer <- x %>%
      mutate(totalt = 1) %>%
      mutate(filter_tmp = as.character(!!as.symbol(filter_var))) %>%
      filter(filter_tmp == filter_val) %>%
      arrange(!!as.symbol(open_ans_var)) %>%
      select(., !!as.symbol(open_ans_var)) %>%
      filter(!!as.symbol(open_ans_var) != "") %>%
      mutate(base_chart = n())

  }

  list_num <- seq(1, 3000, by=72) ## 72 is the number of text fit in a slide

  for (i in 1:length(list_num)) {
    if (dim(open_answer)[1] < list_num[i]) {
      slide_num = i-1
      break

    }
  }

  for (slide in 1:slide_num) {

    if (max(slide_num) > 1) {
      if (slide == 1) {
        start_text = 1
        end_text = list_num[slide+1]

      } else if (slide == max(slide_num)) {
        start_text = list_num[slide] + 1
        end_text = dim(open_answer)[1]

      } else {
        start_text = list_num[slide] + 1
        end_text = list_num[slide+1]

      }

    } else {
      start_text = 1
      end_text = dim(open_answer)[1]

    }

    text_obj <- fpar(ftext("\u2022", fp_text(color = "#e14646", font.size = 10)),
                     ftext("  ", fp_text(color = "#e14646", font.size = 10)),
                     ftext(pull(open_answer, !!as.symbol(open_ans_var))[start_text:end_text], fp_text(color = "#0f283c", font.size = 10)),
                     ftext("\n", fp_text(color = "#e14646", font.size = 10)))

    pptx_obj <- epinion_addslide(pptx_obj, rp_analysis = rp_analysis,
                                 slide_title = slide_title,
                                 slide_subtitle = slide_subtitle,
                                 chart_title = chart_title,
                                 footer = footer,
                                 base_chart = paste0("N=", unique(open_answer[["base_chart"]])),
                                 layout = layout,
                                 text_obj = text_obj)

  }

  pptx_obj

}

# ##############################################################################
# Prepare layout used for report
# ##############################################################################

epinion_addslide = function(pptx_obj, chart_obj = empty_content(), rp_analysis,
                            slide_title = empty_content(),
                            slide_subtitle = empty_content(),
                            chart_title = empty_content(),
                            footer = empty_content(),
                            base_chart = empty_content(),
                            layout,
                            mean_tbl_obj = empty_content(),
                            text_obj = empty_content()) {

  if (grepl("with Mean", rp_analysis))
    # (rp_analysis == "Freq Grid with Mean" |
    #   rp_analysis == "Crosstabs Row with Mean"|
    #   rp_analysis == "Multiple Crosstabs Row (By Case) with Mean")
  {
    # Layout with 1 placeholder for chart and 1 placeholder for mean table
    pptx_obj <- pptx_obj %>%
      add_slide(layout = "7_1 Content Placeholder", master = "Office-tema") %>%
      ph_with(value = chart_obj, location = ph_location_label(ph_label = "Content Placeholder")) %>%
      ph_with(value = mean_tbl_obj, location = ph_location_label(ph_label = "Content Placeholder Mean"),
              header = TRUE, alignment = "c") %>%
      ph_with(value = slide_title, location = ph_location_label(ph_label = "Title")) %>%
      ph_with(value = slide_subtitle, location = ph_location_label(ph_label = "Sub Title")) %>%
      ph_with(value = chart_title, location = ph_location_label(ph_label = "Chart Title")) %>%
      ph_with(value = base_chart, location = ph_location_label(ph_label = "Base Number")) %>%
      # ph_with(value = length(.) - 1, location = ph_location_label(ph_label = "Slide Number Placeholder")) %>%
      ph_with(value = footer, location = ph_location_label(ph_label = "Footer Placeholder")) %>%
      move_slide(index = length(.), to = length(.) - 1)

  } else if (grepl("Open answer",rp_analysis, fixed = TRUE)) {
    # Layout for open answer
    pptx_obj <- pptx_obj %>%
      add_slide(layout = "6_1 Content Placeholder", master = "Office-tema") %>%
      ph_with(value = text_obj, location = ph_location_label(ph_label = "Content Placeholder")) %>%
      ph_with(value = slide_title, location = ph_location_label(ph_label = "Title")) %>%
      ph_with(value = slide_subtitle, location = ph_location_label(ph_label = "Sub Title")) %>%
      ph_with(value = chart_title, location = ph_location_label(ph_label = "Chart Title")) %>%
      ph_with(value = base_chart, location = ph_location_label(ph_label = "Base Number")) %>%
      ph_with(value = footer, location = ph_location_label(ph_label = "Footer Placeholder")) %>%
      move_slide(index = length(.), to = length(.) - 1)

  } else {
    pptx_obj <- pptx_obj %>%
      add_slide(layout = layout, master = "Office-tema") %>%
      ph_with(value = chart_obj, location = ph_location_label(ph_label = "Content Placeholder")) %>%
      ph_with(value = slide_title, location = ph_location_label(ph_label = "Title")) %>%
      ph_with(value = slide_subtitle, location = ph_location_label(ph_label = "Sub Title")) %>%
      ph_with(value = chart_title, location = ph_location_label(ph_label = "Chart Title")) %>%
      ph_with(value = base_chart, location = ph_location_label(ph_label = "Base Number")) %>%
      ph_with(value = footer, location = ph_location_label(ph_label = "Footer Placeholder")) %>%
      move_slide(index = length(.), to = length(.) - 1)

  }

  pptx_obj

}


# ##############################################################################
# Define theme color used for report
# ##############################################################################

theme_10_colors <- list(
  c("#0f283c"),
  c("#0f283c", "#641e3c"),
  c("#0f283c", "#641e3c", "#233ca0"),
  c("#0f283c", "#641e3c", "#233ca0", "#68838b"),
  c("#0f283c", "#641e3c", "#233ca0", "#68838b", "#ba7384"),
  c("#0f283c", "#641e3c", "#233ca0", "#68838b", "#ba7384", "#a7c7d7"),
  c("#0f283c", "#641e3c", "#233ca0", "#68838b", "#ba7384", "#a7c7d7", "#004337"),
  c("#0f283c", "#641e3c", "#233ca0", "#68838b", "#ba7384", "#a7c7d7", "#004337", "#73a89a"),
  c("#0f283c", "#641e3c", "#233ca0", "#68838b", "#ba7384", "#a7c7d7", "#004337", "#73a89a", "#c18022"),
  c("#0f283c", "#641e3c", "#233ca0", "#68838b", "#ba7384", "#a7c7d7", "#004337", "#73a89a", "#c18022", "#ebc882")
)


theme_Scale5_PN99 <- list(
  c("#004337", "#73a89a", "#a7c7d7", "#ebc882", "#c18022", "#e8e1d5")
)

theme_Scale10_NP99 <- list(
  c("#641e3c", "#9c305e", "#ba7384", "#c18022", "#ebc882", "#a7c7d7", "#73a89a", "#68838b", "#233ca0", "#004337", "#0f283c", "#e8e1d5")
)

# ##############################################################################
# Other functions
# ##############################################################################
# Print status of report processing in the console
print_status = function(rp_analysis, row_var, col_var) {
  print(paste0(rp_analysis, " - ",
               row_var, " - ",
               col_var, " - Done!"))

}

# Add percent sign (%) to expss ctables
epinion_add_percent = function(x, digits = get_expss_digits(), excluded_rows = "count", ...) {
  nas = is.na(x)
  x[nas] = ""

  cols_idx = 2:dim(x)[2]

  for (col in cols_idx) {
    for (row in 1:dim(x)[1]) {
      if (!grepl(excluded_rows, x[row, 1], perl = TRUE)) {
        if (suppressWarnings(is.na(as.numeric(as.character(x[row,col]))))) {
          x[row,col] = sub(" ", "% ", trimws(x[row,col]))

        }
        else {
          x[row,col] = paste0(trimws(x[row,col]), "%")

        }
      }
    }
  }
  x <- x[!grepl("Std. dev.", x$row_labels),]
  x <- x[!grepl("Unw. valid N", x$row_labels),]
  x
}

# Format table for PPT
epinion_format_table_PPT = function(mean_tbl, font_size = 10, font_color = "#0f283c") {

  # Set default theme
  mean_table_theme <- set_flextable_defaults(
    font.family = "Arial",
    font.size = font_size,
    font.color = font_color,
    text.align = "c"
  )

  ft <- flextable(mean_tbl) %>%
    fontsize(size = font_size, part = "all") %>%
    border_remove() %>%
    bg(bg = "transparent", part = "all") %>%
    bold(bold = TRUE, part = "header") %>%
    align(align = "center", part = "all") %>%
    fix_border_issues %>%
    height_all(height = 0.95)

  do.call(set_flextable_defaults, mean_table_theme)

  ft
}

# modified 'as_bar_stack' function
# source: mschart::as_bar_stack
epinion_as_bar_stack = function (x, dir = "vertical", percent = FALSE, gap_width = 50) {
  stopifnot(inherits(x, "ms_barchart"))
  grouping <- "stacked"

  if (percent)
    grouping <- "percentStacked"
  x <- chart_settings(x, grouping = grouping, dir = dir, gap_width = gap_width,
                      overlap = 100)

  # x <- chart_data_stroke(x, values = "transparent")
  if (dir == "horizontal")
    x <- chart_theme(x = x, title_x_rot = 270, title_y_rot = 0)
  x

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

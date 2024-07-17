#' Read data input
#'
#' This function reads different type of data input (xlsx, csv, rds, sav, txt, sas) into a dataframe.
#' It also allows users to choose only the required variables for data processing.
#'
#' @param file the name of the file which the data are to be read from.
#' @param header a logical value indicating whether the file contains the names of the variables as its first line.
#' @param sep the field separator character.
#' @param fileEncoding character string declares the encoding used on a file.
#' @param ... Additional arguments may apply.
#' @return A dataframe of data input
#' @examples
#' data_test <- epinion_read_data(file = example("sample_test.xlsx"), cols = c("bagg2_rc", "bagg3_rc", "kommum_03"))
#' data_test_sav <- epinion_read_data(file = example("sample_test.sav"))

epinion_read_data <- function(file, header = TRUE, sep = ";", fileEncoding = "UTF-16LE" , cols = everything(), ...){

  if (!file.exists(file)) {
    message("The file does not exist")
  } else {

    extension<-tools::file_ext(file)

    if (extension == "xlsx") {
      df = openxlsx::read.xlsx(xlsxFile = file, colNames = TRUE)
    }

    if (extension == "csv") {
      df = read.csv(file = file, header = header, sep = sep)
    }

    if (extension == "rds") {
      df = readRDS(file = file)
    }

    if (extension == "sav") {
      df = haven::read_sav(file = file)
    }

    if (extension == "sas7bdat") {
      df = haven::read_sas(data_file = file)
    }

    if (extension == "txt") {
      df = read.table(file = file, header = header,
                      quote = "", sep = sep,
                      fileEncoding = fileEncoding,
                      skipNul = TRUE)
    }

    require(dplyr)

    df <- df %>%
      select(all_of(cols))

    return(df)

  }
}

validate_weight_matrix <- function(df, df_weight_matrix) {
  # Check variable
  for (var in unique(df_weight_matrix$weight_var)) {
    if (!var %in% names(df)) {
      stop("Variable '", var, "' does not exist in the data. Please check!")

    }
  }

  # Check value
  for (idx in 1:dim(df_weight_matrix)[1]) {
    if (!df_weight_matrix$weight_cat[idx] %in% df[[df_weight_matrix$weight_var[idx]]]) {
      stop(paste0("There is no case in the data with value ", df_weight_matrix$weight_cat[idx], " at variable '",
                  df_weight_matrix$weight_var[idx], "'. Please check the weight matrix!"))
    }

  }

  # Check weight final
  if (length(unique(df_weight_matrix$weight_final)) > 1) {
    stop("The name of the weight final is not unique. Please check the weight matrix!")

  }

  # Check sum of each weight dim
  require(dplyr)

  df_weight_matrix <- df_weight_matrix %>%
    group_by(weight_var) %>%
    mutate(weight_total = sum(weight_value, na.rm = TRUE)) %>%
    select(., weight_var, weight_total) %>%
    unique()

  for (idx in 1:dim(df_weight_matrix)[1]) {
    if (df_weight_matrix$weight_total[idx] != 1) {
      stop("The sum weight value at variable '", df_weight_matrix$weight_var[idx],
           "' does not equal to 1. Please check the weight matrix!")

    }

  }

}


#' Weigh data
#'
#' This function used to weigh the data so that it could meet the distribution of population's demographics.
#'
#' @param data_input data to be weighted (dataframe).
#' @param weight_matrix_file weight matrix (.txt file).
#' @param unique_id_var variable used to identify each case as a unique in the data.
#' @param fileEncoding character string declares the encoding used on the weight_matrix_file.
#' @return A data with weight variable included
#' @examples
#' data_test_sav <- epinion_read_data(file = example("sample_test.sav")) %>%
#' as.data.frame()
#' data_weighted <- epinion_weighting_tool(data_input = data_test_sav,
#' weight_matrix_file = example("Weight matrix.txt"),
#' unique_id_var = "RespId")

epinion_weighting_tool <- function(data_input, weight_matrix_file,
                                   unique_id_var,
                                   fileEncoding = "UTF-8") {

  if (!require("pacman")) install.packages("pacman")
  pacman::p_load(dplyr, anesrake)

  # Get data and weight matrix
  df <- data_input

  weight_matrix <- epinion_read_data(weight_matrix_file,
                                     header = FALSE,
                                     sep = ";",
                                     fileEncoding = fileEncoding) %>%
    rename(weight_var = V1,
           weight_cat = V2,
           weight_value = V3,
           weight_final = V4) %>%
    mutate(weight_value = as.numeric(weight_value))

  # Validate weight matrix
  validate_weight_matrix(df, weight_matrix)

  # Create weight dimensions in the data based on weight matrix
  weight_dim = unique(weight_matrix$weight_var)
  weight_tmp = c("dim1", "dim2", "dim3", "dim4", "dim5",
                 "dim6", "dim7", "dim8", "dim9", "dim10",
                 "dim11", "dim12", "dim13", "dim14", "dim15")

  targets = list()

  for (var_idx in 1:length(weight_dim)) {
    # 1. Convert weighting dim variables in the data to factors
    df <- df %>%
      mutate(!!as.symbol(weight_tmp[var_idx]) := as.factor(!!as.symbol(weight_dim[var_idx])))

    # 2. Define weight dimensions
    dim_tmp = filter(weight_matrix, weight_var == weight_dim[var_idx])$weight_value
    names(dim_tmp) <- filter(weight_matrix, weight_var == weight_dim[var_idx])$weight_cat

    # 3. Create weight matrix.
    targets = append(targets, list(dim_tmp))

  }

  # Important: to use the same variable names in the dataset
  names(targets) <- weight_tmp[1:length(weight_dim)]

  # 4. Weight
  anesrakefinder(targets, df, choosemethod = "total")

  outsave <- anesrake(inputter     = targets,
                      dataframe    = df,
                      caseid       = df[[unique_id_var]],
                      choosemethod = "total", # c("total", "max", "average", "totalsquared", "maxsquared", "averagesquared")
                      type         = "pctlim",
                      pctlim       = 0.00001)

  print(summary(outsave))

  # 5. Add weights to the dataset
  df[[unique(weight_matrix$weight_final)]]  <- unlist(outsave[1])

  df

}

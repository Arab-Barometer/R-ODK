#' Randomize Choice Order
#'
#' @description
#' Given a single set of choices already programmed on an xlsx sheet in ODK format,
#' this function will create an xlsx sheet containing all the possible permutations
#' of the choice order with any choice of value 90 or higher frozen at the end of
#' the order by default. It is also possible to freeze choice orders at the start
#' of the list, or at the start and end.
#'
#' It is then possible to directly copy and paste the output
#' into the XLSForm, but it is recommended to upload this xlsx form to KoBoToolbox
#' and use either `select_one_from_file` or `select_multiple_from_file` question
#' type and reference the output file.
#'
#' **These questions will still require a calculate type question in the survey sheet to determine which set of choices to display.**
#' The user can copy the calculation printed out by the function in a warning message and paste it
#' into the **calculation** column for a **calculate** type question.
#'
#' @details
#' The user must provide the **number of rows** they want to freeze at the top and/or
#' bottom of the choice list. For example, if there are seven options but the user
#' wants the first item to always appear first and the last two items to always appear
#' last. The `freeze_top` parameter should then be set to `1` and the `freeze_bottom`
#' parameter should be set to `2`.
#'
#' @param choice_sheet A (path to a) ODK xlsx cohice sheet that has a set of programmed choices
#' @param freeze_top Optional - Number of rows at the top of the choice list that should stay frozen. Defaults to NULL
#' @param freeze_bottom Optional - Number of rows at the bottom of the choice list that should stay froze. Defaults to NULL. If both `freeze_top` and `freeze_bottom` are NULL, items with a value less than 90 will be randomized.
#' @param pth Optional - A file path to where the output should be saved. Defaults to the current working directory
#' @param output Optional - If TRUE will return the xlsx form as a data frame to R. If FALSE the only output is the xlsx file. Defaults to TRUE
#'
#' @return A ODK-style choice sheet as a data frame and an .xlsx sheet
#' @export
randomize_choice_order <- function(choice_sheet,
                                   freeze_top = NULL,
                                   freeze_bottom = NULL,
                                   pth = NULL,
                                   output = TRUE){

  choice_sheet <- readxl::read_xlsx(choice_sheet)
  n_rows <- nrow(choice_sheet)

  if(is.null(freeze_top) & is.null(freeze_bottom)){

    freeze_rows <- choice_sheet[choice_sheet$name >= 90,]
    randomize_rows <- choice_sheet[choice_sheet$name < 90,]

  } else if(is.null(freeze_top) & !is.null(freeze_bottom)){

    freeze_rows <- choice_sheet[(n_rows - freeze_bottom + 1):n_rows,]
    randomize_rows <- choice_sheet[1:(n_rows + freeze_bottom),]

  } else if(!is.null(freeze_top) & is.null(freeze_bottom)){

    freeze_rows <- choice_sheet[1:freeze_top,]
    randomize_rows <- choice_sheet[(freeze_top + 1):n_rows,]

  } else {
    top_rows <- choice_sheet[1:freeze_top,]
    bottom_rows <- choice_sheet[(n_rows - freeze_bottom + 1):n_rows,]
    randomize_rows <- choice_sheet[(freeze_top+1):(n_rows - freeze_bottom),]
  }

  all.perm <- combinat::permn(1:nrow(randomize_rows))

  randomized_top <- function(x){
    new_df <- rbind(randomize_rows[order(all.perm[[x]]),],
                    freeze_rows)
    new_df["RandomChoice"] <- x
    new_df
  }

  randomized_bottom <- function(x){
    new_df <- rbind(freeze_rows,
                    randomize_rows[order(all.perm[[x]]),])
    new_df["RandomChoice"] <- x
    new_df
  }

  randomized_split <- function(x){
    new_df <- rbind(top_rows,
                    randomize_rows[order(all.perm[[x]]),],
                    bottom_rows)
    new_df["RandomChoice"] <- x
    new_df
  }

  if(!is.null(freeze_top) & !is.null(freeze_bottom)){

    randomized_choices <- purrr::map_df(1:length(all.perm),
                                        randomized_split)

  } else if(!is.null(freeze_top) & is.null(freeze_bottom)){

    randomized_choices <- purrr::map_df(1:length(all.perm),
                                        randomized_bottom)

  } else {

    randomized_choices <- purrr::map_df(1:length(all.perm),
                                        randomized_top)
  }


  insight::format_warning("Do not forget to create a calculate question in the {.b survey} sheet with the following calculation:",
                          stringr::str_glue("round((once(random())*{length(all.perm) - 1}) + 1, 0)"))


  if(!is.null(pth)){
    if(stringr::str_detect(pth,"/",negate = T)){
      pth <- stringr::str_glue("{pth}/")
    }
  }

  df_name <- ifelse(is.null(pth),
                    paste0(unique(choice_sheet[["list_name"]]),"_randomized"),
                    paste0(pth,unique(choice_sheet[["list_name"]]),"_randomized"))

  writexl::write_xlsx(randomized_choices,
                      path = paste0(df_name,".xlsx"))


  if(output){randomized_choices}

}

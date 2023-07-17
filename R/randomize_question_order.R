#' Randomize Question Order
#'
#' @description
#' Given a single set of grid questions already programmed on an .xlsx sheet with ODK columns,
#' this function will create a new .xlsx sheet with the questions randomized in every possible
#' order. The output captured in the new .xlsx file can be copied and pasted directly into
#' an XLSForm for ODK.
#'
#' To keep randomization to a minimum, questions that are never shown together
#' (e.g., one is only shown to group A while another is only shown to group B) can
#' be grouped together using the `.group_id` parameter. Questions that are in the
#' same group will be treated as one question, reducing the number of permutations.
#' See details for a more thorough explanation.
#'
#' @details
#' To produce question randomization in ODK, this function takes a set of questions that are
#' grouped together, then creates a new group for every permutation of the questions. Each
#' permutation is assigned a number 1 to N where N is the number of possible permutations.
#' A calculate row is added to the beginning of the programming. In that row, there is a
#' calculation that will randomly produce a number 1 to N. A set of questions will appear
#' based on the random number.
#'
#' The `.group_id` parameter can be used to reduce the number of permutations by grouping
#' certain questions together. For example, say there is a group of four questions that
#' need to be randomized, but two questions are only shown to group A while two questions
#' are only shown to group B. By grouping together an A question and a B question, the number
#' of permutations can be reduced from 24 (4P4) to 2 (2P2).
#'
#' The `.group_id` parameter takes a vector. To use it, first group the questions together.
#' Continuing with the same example as before, call our four questions A1, B1, A2, B2. The
#' A questions are only shown to group A and the B questions are only shown to group B.
#' Say A1 and B1 are grouped together in group 1 and A2 and B2 are grouped together in
#' group 2. For the `.group_id` parameter, use the vector `c(1,1,2,2)`. Each element
#' in the vector is a group assignment. The first two questions are A1 and B1, which
#' are both assigned 1. The last two questions are A2 and B2 which are both assigned 2.
#' The `.group_id` vector must be the same length as the number of questions that are
#' being randomized.
#'
#' @param xlsform A (path to a) ODK xlsx survey sheet that has a set of programmed grid questions
#' @param .group_id Optional - A vector of either integers or characters that describe how the questions should be grouped together. Defaults to NULL, which randomized every question
#' @param pth Optional - A file path to where the output should be saved. Defaults to the current working directory
#' @param output Optional - If TRUE will return the xlsx form as a data frame to R. If FALSE the only output is the xlsx file. Defaults to TRUE
#'
#' @return A ODK-style survey sheet as a data frame and an .xlsx sheet
#' @export
randomize_question_order <- function(xlsform,
                                     .group_id = NULL,
                                     pth = NULL,
                                     output = TRUE){

  odk_grid <- readxl::read_xlsx(xlsform)

  if(!all(c("name","relevant") %in% names(odk_grid))){
    stop("Make sure the form you are loading is an ODK survey sheet.")
  }

  # Extracting question name:
  q_name <- stringr::str_extract(odk_grid[["name"]][1],"[:alnum:]+") ## Question number

  # Breaking df into components:
  start_group_row <- dplyr::slice_head(odk_grid) ## First row
  shuffle_rows <- dplyr::slice(odk_grid,-c(1,nrow(odk_grid))) ## Row to create permutations with
  end_group_row <- dplyr::slice_tail(odk_grid) ## Last row

  # Creating calculate question name:
  calc_name <- paste0(q_name,"_rand")

  # Number of rows:
  m <- dim(shuffle_rows)[1]

  # Add new column to ID group:
  if(!is.null(.group_id)) {
    if (length(unique(.group_id)) > 1){
      shuffle_rows$group_id <- .group_id
    }else{
      insight::format_warning(
        "You had put all the questions in the same group.",
        "This means there is nothing to randomize. In case you meant to randomize everything, this function produced a xlsx file will all rows randomized."
      )
      shuffle_rows$group_id <- 1:m
    }

  }else{
    shuffle_rows$group_id <- 1:m
  }

  # Break tibble apart by groups:
  loose_groups <- split(shuffle_rows,
                        shuffle_rows[["group_id"]])

  # Creating permutations:
  all.perm <- combinat::permn(names(loose_groups))

  for (i in 1:length(all.perm)) {

    full_countraint <- paste0(calc_name, " = ", i)

    new_relevance <- ifelse(is.na(start_group_row[["relevant"]]),
                            full_countraint,
                            paste0(start_group_row[["relevant"]]," and ",full_countraint))


    new_start_group_row <- dplyr::mutate(start_group_row,
                                         name = replace(name, 1,
                                                        paste0(q_name,"_grid_",i)),
                                         relevant = replace(relevant, 1,
                                                            new_relevance) ### Close replace()
      )

    if (i == 1) {
      .new_df <- loose_groups[order(all.perm[[i]])]
      .new_df <- dplyr::bind_rows(.new_df)

      # Adding start and end group
      .new_df <- dplyr::add_row(.new_df,
                         new_start_group_row,
                         .before = i)
      .new_df <- dplyr::add_row(.new_df,
                         end_group_row,
                         .after = m + 1)

    } else {
      .new_order <- loose_groups[order(all.perm[[i]])]
      .new_order <- dplyr::bind_rows(.new_order)

      # Adding start and end group
      .new_order <- dplyr::add_row(.new_order,
                            new_start_group_row,
                            .before = 1)
      .new_order <- dplyr::add_row(.new_order,
                            end_group_row,
                            .after = m + 1)

      # Combining all the data
      .new_df <- rbind(.new_df,.new_order)
    }

  }


  # Adding calculate question:
  .new_df <- dplyr::add_row(.new_df,
                     type = "calculate",
                     name = calc_name,
                     calculation = stringr::str_glue("once(round((random()*{length(all.perm) - 1})+1,0))"),
                     .before = 1)

  .new_df$group_id <- NULL


  if(!is.null(pth)){
    if(stringr::str_detect(pth,"/",negate = T)){
      pth <- stringr::str_glue("{pth}/")
    }
  }

  df_name <- ifelse(is.null(pth),
                    paste0(q_name,"_randomized"),
                    paste0(pth,q_name,"_randomized"))

  writexl::write_xlsx(.new_df,
                      path = paste0(df_name,".xlsx"))


  if(output){.new_df}
}


# library(combinat)
# library(readxl)
# library(dplyr)
# library(purrr)
# library(randomizr)

group_permutations <- function(.df,
                               .group_id,
                               pth = NULL){

  # Extracting question name:
  q_name <- stringr::str_extract(.df[["name"]][1],"[:alnum:]+") ## Question number

  # Breaking df into components:
  start_group_row <- dplyr::slice_head(.df) ## First row
  shuffle_rows <- dplyr::slice(.df,-c(1,nrow(.df))) ## Row to create permutations with
  end_group_row <- dplyr::slice_tail(.df) ## Last row

  # Creating calculate question name:
  calc_name <- paste0(q_name,"_rand")


  # Number of groups:
  n <- length(unique(.group_id))
  # Number of rows:
  m <- length(.group_id)

  # Add new column to ID group:
  shuffle_rows$group_id <- .group_id

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
                                         name = replace(.$name, 1,
                                                        paste0(q_name,"_grid_",i)),
                                         relevant = replace(.$relevant, 1,
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
                     calculation = stringr::str_glue("once(random()) * {length(all.perm) - 1}
                                                     + 1"),
                     .before = 1)

  .new_df$group_id <- NULL



  df_name <- ifelse(is.null(pth),
                    paste0(q_name,"_randomized"),
                    paste0(pth,q_name,"_randomized"))

  writexl::write_xlsx(.new_df,
                      path = paste0(df_name,".xlsx"))


  return(.new_df)
}

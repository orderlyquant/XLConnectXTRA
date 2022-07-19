
load_xl_template <- function(template_name) {
  wb <- loadWorkbook(template_name)
  setStyleAction(wb, XLC$"STYLE_ACTION.NONE")

  # A template workbook is composed of template sheets. Account for them
  # on loading template, so they can be removed at end of filling the template.
  templates <- getSheets(wb)

  return(
    list(
      wb = wb,
      templates = templates
    )
  )
}


duplicate_template_sheet <- function(wb, template_name, uniq_name) {
  cloneSheet(wb, sheet = template_name, name = uniq_name)

  named_ranges <- getReferenceFormula(wb, getDefinedNames(wb))
  named_ranges <- named_ranges[
    str_detect(named_ranges, template_name)
  ]

  original_names <- names(named_ranges)

  named_ranges <- str_replace(named_ranges, "^.+!", "")
  names(named_ranges) <- str_replace(
    original_names,
    str_c(template_name, "_", sep = ""),
    ""
  )

  # Define ranges for newly created sheet
  for (i in 1:length(named_ranges)) {
    createName(
      wb,
      name = ifelse(
        nchar(uniq_name) > 0,
        str_c(uniq_name, names(named_ranges)[i], sep = "_"),
        str_c(uniq_name, names(named_ranges)[i], sep = "_")
      ),
      formula = str_c(str_c("'", uniq_name, "'", sep = ""), named_ranges[i], sep = "!")
    )
  }
}



get_named_ranges_for_sheet <- function(wb, sheet_name) {
  named_ranges <- getReferenceFormula(wb, getDefinedNames(wb))
  named_ranges <- named_ranges[
    str_detect(named_ranges, sheet_name)
  ]

  return(named_ranges)
}



cleanup_template <- function(wb, templates) {
  for (i in 1:length(templates)) {
    removeSheet(wb, templates[i])

    named_ranges <- getReferenceFormula(wb, getDefinedNames(wb))
    named_ranges <- named_ranges[
      str_detect(named_ranges, templates[i])
    ]

    for (i in 1:length(named_ranges)) {
      removeName(wb, names(named_ranges)[i])
    }
  }
}



insert_columns <- function(df, column_content, positions) {
  positions <- positions[order(positions)]

  for (i in 1:length(positions)) {
    position <- positions[i]

    n.orig.cols <- NCOL(df)

    # position can't be greater than number of columns + 1
    position <- ifelse(position > (n.orig.cols + 1), (n.orig.cols + 1), position)

    # Determine column name. The convention will be Insert.01, Insert.02, etc.
    # Check for previously inserted columns.
    current.cols <- names(df)
    num.prev.inserted.cols <- sum(str_detect(current.cols, "^Insert.[0-9]+"), na.rm = TRUE)
    new.col.name <- str_c("Insert", sprintf("%02d", num.prev.inserted.cols + 1), sep = ".")

    # Set up final column order
    begin.cols <- ifelse((1:n.orig.cols) < position, current.cols, NA)
    end.cols <- ifelse((1:n.orig.cols) >= position, current.cols, NA)

    new.col.order <- c(begin.cols, new.col.name, end.cols)
    new.col.order <- new.col.order[complete.cases(new.col.order)]

    # Add column
    df[, new.col.name] <- column_content

    # Put new column in proper order

    df <- df[, new.col.order]
  }

  return(df)
}

insert_rows <- function(df, before_rows, row_content = NA, repeat_row_content = FALSE) {
  blank_row <- df[1, ]
  blank_row[1, ] <- NA

  if (repeat_row_content) {
    blank_row[1, ] <- row_content
  } else {
    blank_row[1, 1] <- row_content
  }

  before_rows <- before_rows[order(before_rows)]

  increment_offset <- 0

  for (i in 1:length(before_rows)) {
    position <- before_rows[i] + increment_offset

    n_orig_rows <- NROW(df)

    # handle position = 1 separately
    if (position <= 1) {
      position <- max(position, 1)
      df <- rbind(
        blank_row,
        df[position:nrow(df), ]
      )
    } else if (position > n_orig_rows) {
      df <- rbind(
        df,
        blank_row
      )
    } else {
      df <- rbind(
        df[1:(position - 1), ],
        blank_row,
        df[position:nrow(df), ]
      )
    }

    increment_offset <- increment_offset + 1
  }

  return(df)
}


write_named_region_safe <- function(file, data, name, nrows = 5000) {
  header <- TRUE
  cat(str_c("          rows left in df: ", nrow(data), "\n"))

  while (nrow(data) > nrows) {
    writeNamedRegionToFile(
      file = file,
      data = data[1:nrows, ],
      name = name,
      header = header,
      styleAction = XLC$"STYLE_ACTION.NONE"
    )
    # should only be true once
    header <- FALSE

    data <- data[(nrows + 1):nrow(data), ]
    cat(str_c("          rows left in df: ", nrow(data), "\n"))
  }

  writeNamedRegionToFile(
    file = file,
    data = data,
    name = name,
    header = header,
    styleAction = XLC$"STYLE_ACTION.NONE"
  )
}

append_named_region_safe <- function(object, data, name, nrows = 5000) {
  header <- TRUE
  cat(str_c("          rows left in df: ", nrow(data), "\n"))

  while (nrow(data) > nrows) {
    appendNamedRegion(
      object = object,
      data = data[1:nrows, ],
      name = name,
      header = header
    )
    # should only be true once
    header <- FALSE

    data <- data[(nrows + 1):nrow(data), ]
    cat(str_c("          rows left in df: ", nrow(data), "\n"))
  }

  appendNamedRegion(
    object = object,
    data = data,
    name = name,
    header = header
  )
}

update_report_names <- function(lst, prefix, sep = "_") {
  for (i in 1:length(lst)) {
    rpt_name <- attr(lst[[i]], "name")
    new_rpt_name <- paste(prefix, rpt_name, sep = sep)

    attr(lst[[i]], "name") <- new_rpt_name
  }

  return(lst)
}

write_report_object <- function(xl_obj, data_obj) {
  setStyleAction(xl_obj, XLC$"STYLE_ACTION.NONE")

  obj_type <- attr(data_obj, "type")

  if (obj_type == "data.frame") {
    writeNamedRegion(
      xl_obj,
      data_obj,
      attr(data_obj, "name"),
      header = attr(data_obj, "incl_header")
    )
  }

  if (obj_type == "ggplot") {
    fig_dims <- attr(data_obj, "dims")

    tmp_fig_file <- fs::file_temp(
      tmp_dir = here::here(), ext = "png"
    ) %>% as.character()

    ggsave(
      tmp_fig_file, data_obj,
      height = fig_dims[1], width = fig_dims[2], dpi = 288,
      bg = "white"
    )
    addImage(xl_obj, tmp_fig_file, attr(data_obj, "name"), originalSize = FALSE)

    # remove temp file
    fs::file_delete(tmp_fig_file)
  }

  # TODO: improve this logic
  # If execution got to here, data_obj wasn't a configured type

  if (!(obj_type %in% c("data.frame", "ggplot"))) {
    writeNamedRegion(
      xl_obj,
      data_obj,
      attr(data_obj, "name"),
      header = attr(data_obj, "incl_header")
    )
  }
}

report_reportables <- function(xl_obj, tpl_name, instance_name, rpt_list) {
  duplicate_template_sheet(xl_obj, tpl_name, instance_name)

  # update names in rpt_list
  rpt_list <- update_report_names(rpt_list, instance_name)

  for (i in 1:length(rpt_list)) {
    write_report_object(xl_obj, rpt_list[[i]])
  }
}

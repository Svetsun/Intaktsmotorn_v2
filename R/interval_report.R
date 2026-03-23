# ===================== INTERVALLRAPPORT =====================
# Temporal resolution helpers and main build function.
# Sourced at top level of app.r so all shared helpers are in scope.

# Returns the timpris to use for a (consultant, uppdrag, uppgift) combination
# in the month identified by target_first_day.
# Rule: a TimprisHistory change recorded in month X is active from month X.
#       → keep rows where created_at is before first day of month X+1.
# Fallback: current Uppgift$timpris via lookup_timpris_from_uppgift().
resolve_timpris_for_month <- function(timpris_history, uppgift_df,
                                       consultant_id, uppdrag_id, uppgift_id,
                                       target_first_day) {
  .cid <- as.character(consultant_id)
  .uid <- as.character(uppdrag_id)
  .tid <- as.character(uppgift_id)
  .next_month_first <- as.Date(target_first_day) %m+% months(1)

  th <- sanitize_nulls_df(timpris_history) |> coerce_dates(c("created_at"))

  if (nrow(th) > 0 &&
      all(c("consultant_id", "uppdrag_id", "uppgift_id", "timpris", "created_at") %in% names(th))) {
    th$consultant_id <- as.character(th$consultant_id)
    th$uppdrag_id    <- as.character(th$uppdrag_id)
    th$uppgift_id    <- as.character(th$uppgift_id)
    th$timpris       <- suppressWarnings(as.numeric(th$timpris))

    mask <- th$consultant_id == .cid &
            th$uppdrag_id    == .uid &
            th$uppgift_id    == .tid &
            !is.na(th$created_at) &
            th$created_at    <  .next_month_first &
            !is.na(th$timpris)

    hit <- th[mask, , drop = FALSE]
    if (nrow(hit) > 0) {
      hit <- hit[order(hit$created_at, decreasing = TRUE), , drop = FALSE]
      return(as.numeric(hit$timpris[1]))
    }
  }

  lookup_timpris_from_uppgift(uppgift_df,
                               consultant_id = .cid,
                               uppdrag_id    = .uid,
                               uppgift_id    = .tid)
}

# Returns the grundlön to use for a consultant in the year target_year.
#
# Case 1 – New consultant (startdatum falls in target_year or later):
#   Their register grundlön IS their starting salary.  Return it immediately;
#   any GrundlonHistory entries recorded in the same year apply from Jan 1
#   of the NEXT year (handled automatically by Case 2 in subsequent calls).
#
# Case 2 – Existing consultant (startdatum before target_year):
#   Yearly activation rule: a change recorded in year Y is active from Y+1.
#   → keep only rows where year(created_at) < target_year, pick the most recent.
#   Fallback: current Konsulter$grundlon, or 0 if missing.
resolve_grundlon_for_month <- function(grundlon_history, konsulter_df,
                                        consultant_id, target_first_day) {
  .cid  <- as.character(consultant_id)
  .fd   <- as.Date(target_first_day)
  .year <- as.integer(format(.fd, "%Y"))

  # Resolve register row once — needed for both cases and as fallback.
  kons <- sanitize_nulls_df(konsulter_df)
  kons$consultant_id <- as.character(kons$consultant_id)
  kons$grundlon      <- suppressWarnings(as.numeric(kons$grundlon))
  krow <- kons[kons$consultant_id == .cid, , drop = FALSE]
  reg_grundlon <- if (nrow(krow) > 0 && !is.na(krow$grundlon[1]))
    as.numeric(krow$grundlon[1]) else 0

  # Never count salary before actual employment start.
  if (nrow(krow) > 0 && "startdatum" %in% names(krow)) {
    sd <- suppressWarnings(as.Date(as.character(krow$startdatum[1])))
    if (!is.na(sd) && !is.na(.fd)) {
      ld <- as.Date(month_bounds(.year, as.integer(format(.fd, "%m")))$end)
      if (ld < sd) return(0)
    }
  }

  # Yearly activation rule:
  # - change recorded in January -> active from Jan 1 same year
  # - change recorded Feb-Dec  -> active from Jan 1 next year
  # For multiple changes mapping to the same active year, latest created_at wins.
  gl <- sanitize_nulls_df(grundlon_history) |> coerce_dates(c("created_at"))

  if (nrow(gl) > 0 &&
      all(c("consultant_id", "grundlon", "created_at") %in% names(gl))) {
    gl$consultant_id <- as.character(gl$consultant_id)
    gl$grundlon      <- suppressWarnings(as.numeric(gl$grundlon))
    gl$active_year   <- ifelse(
      lubridate::month(gl$created_at) == 1,
      as.integer(format(gl$created_at, "%Y")),
      as.integer(format(gl$created_at, "%Y")) + 1L
    )

    mask <- gl$consultant_id == .cid &
            !is.na(gl$created_at) &
            !is.na(gl$active_year) &
            gl$active_year <= .year &
            !is.na(gl$grundlon)

    hit <- gl[mask, , drop = FALSE]
    if (nrow(hit) > 0) {
      hit <- hit[order(hit$active_year, hit$created_at, decreasing = TRUE), , drop = FALSE]
      return(as.numeric(hit$grundlon[1]))
    }
  }

  reg_grundlon
}

# Returns the bonus_grund (%) to use for a consultant in the month at target_first_day.
# Rule: same activation as timpris — change in month X is active from month X+1.
# Fallback: current Konsulter$bonus_grund, or 0 if missing.
resolve_bonus_for_month <- function(bonus_history, konsulter_df,
                                     consultant_id, target_first_day) {
  .cid <- as.character(consultant_id)

  bh <- sanitize_nulls_df(bonus_history) |> coerce_dates(c("created_at"))

  if (nrow(bh) > 0 &&
      all(c("consultant_id", "bonus_grund", "created_at") %in% names(bh))) {
    bh$consultant_id <- as.character(bh$consultant_id)
    bh$bonus_grund   <- suppressWarnings(as.numeric(bh$bonus_grund))

    mask <- bh$consultant_id == .cid &
            !is.na(bh$created_at) &
            bh$created_at    <  target_first_day &
            !is.na(bh$bonus_grund)

    hit <- bh[mask, , drop = FALSE]
    if (nrow(hit) > 0) {
      hit <- hit[order(hit$created_at, decreasing = TRUE), , drop = FALSE]
      return(as.numeric(hit$bonus_grund[1]))
    }
  }

  kons <- sanitize_nulls_df(konsulter_df)
  kons$consultant_id <- as.character(kons$consultant_id)
  kons$bonus_grund   <- suppressWarnings(as.numeric(kons$bonus_grund))
  row <- kons[kons$consultant_id == .cid, , drop = FALSE]
  if (nrow(row) > 0 && !is.na(row$bonus_grund[1])) as.numeric(row$bonus_grund[1]) else 0
}

# ===================== MAIN BUILD FUNCTION =====================
# Returns list(detail, summary, totals).
# - detail:  one row per (month, consultant, customer)
# - summary: one row per (month, consultant) with resolved grundlön/bonus
# - totals:  one row per consultant, summed across the whole interval
build_interval_report <- function(wb, labels,
                                   start_year, start_month,
                                   end_year,   end_month,
                                   bonus_threshold = BONUS_THRESHOLD) {

  start_year  <- as.integer(start_year)
  start_month <- as.integer(start_month)
  end_year    <- as.integer(end_year)
  end_month   <- as.integer(end_month)

  make_empty <- function() {
    list(
      detail = data.frame(
        ar = integer(0), manad = integer(0),
        fornamn = character(0), efternamn = character(0),
        kundnamn = character(0), uppdrag_label = character(0),
        timmar = numeric(0), timpris_used = numeric(0),
        total_summa = numeric(0), debiteringsgrad = numeric(0),
        stringsAsFactors = FALSE
      ),
      summary = data.frame(
        ar = integer(0), manad = integer(0),
        fornamn = character(0), efternamn = character(0),
        total_debiterad = numeric(0),
        grundlon_used = numeric(0), bonus_percent_used = numeric(0),
        bonus_belopp = numeric(0), lon_total = numeric(0),
        debiteringsgrad = numeric(0),
        stringsAsFactors = FALSE
      ),
      totals = data.frame(
        fornamn = character(0), efternamn = character(0),
        total_debiterad = numeric(0),
        grundlon_total = numeric(0), bonus_belopp_total = numeric(0),
        lon_total = numeric(0), debiteringsgrad = numeric(0),
        stringsAsFactors = FALSE
      ),
      arb_hours_by_month = list()
    )
  }

  if (anyNA(c(start_year, start_month, end_year, end_month))) return(make_empty())
  if (start_month < 1 || start_month > 12 || end_month < 1 || end_month > 12) return(make_empty())

  start_first <- as.Date(sprintf("%04d-%02d-01", start_year, start_month))
  end_first   <- as.Date(sprintf("%04d-%02d-01", end_year,   end_month))
  if (start_first > end_first) return(make_empty())

  month_firsts <- seq(start_first, end_first, by = "month")

  # Pre-load all source tables once
  tid     <- sanitize_nulls_df(wb[["Tidrapportering"]]) |>
               coerce_dates(c("startdatum", "slutdatum", "created_at"))
  uppgift <- sanitize_nulls_df(wb[["Uppgift"]]) |> ensure_uppgift_timpris()
  kons    <- sanitize_nulls_df(wb[["Konsulter"]])
  kunder  <- sanitize_nulls_df(wb[["Kunder"]])
  arb     <- sanitize_nulls_df(wb[["ArbetstimmarGrund"]]) |> coerce_dates(c("date"))

  th_hist <- if (!is.null(wb[["TimprisHistory"]]))
    sanitize_nulls_df(wb[["TimprisHistory"]]) |> coerce_dates(c("created_at"))
  else data.frame()

  gl_hist <- if (!is.null(wb[["GrundlonHistory"]]))
    sanitize_nulls_df(wb[["GrundlonHistory"]]) |> coerce_dates(c("created_at"))
  else data.frame()

  bh_hist <- if (!is.null(wb[["BonusHistory"]]))
    sanitize_nulls_df(wb[["BonusHistory"]]) |> coerce_dates(c("created_at"))
  else data.frame()

  if (nrow(tid) == 0) return(make_empty())

  tid <- tid |> mutate(
    consultant_id = as.character(consultant_id),
    uppdrag_id    = as.character(uppdrag_id),
    uppgift_id    = as.character(uppgift_id),
    customer_id   = as.character(customer_id),
    timmar        = suppressWarnings(as.numeric(timmar))
  )

  kons_names <- kons |>
    mutate(consultant_id = as.character(consultant_id),
           fornamn       = as.character(fornamn),
           efternamn     = as.character(efternamn)) |>
    select(consultant_id, fornamn, efternamn)

  kund_map <- kunder |>
    mutate(
      customer_id = as.character(customer_id),
      kundnamn    = ifelse(is.na(customer_namn) | customer_namn == "",
                           customer_id, as.character(customer_namn))
    ) |>
    select(customer_id, kundnamn)

  uppdrag_map <- sanitize_nulls_df(wb[["Uppdrag"]]) |>
    mutate(
      uppdrag_id    = as.character(uppdrag_id),
      uppdrag_label = ifelse(is.na(uppdrag_name) | trimws(as.character(uppdrag_name)) == "",
                             uppdrag_id, as.character(uppdrag_name))
    ) |>
    select(uppdrag_id, uppdrag_label)

  monthly_detail  <- vector("list", length(month_firsts))
  monthly_summary <- vector("list", length(month_firsts))
  arb_hours_list  <- list()

  for (i in seq_along(month_firsts)) {
    fd <- month_firsts[[i]]
    y  <- as.integer(format(fd, "%Y"))
    m  <- as.integer(format(fd, "%m"))
    ld <- as.Date(month_bounds(y, m)$end)

    # Overlap filter: startdatum <= last_day AND slutdatum >= first_day
    tid_m <- tid[!is.na(tid$startdatum) & !is.na(tid$slutdatum) &
                   tid$startdatum <= ld & tid$slutdatum >= fd, , drop = FALSE]
    if (nrow(tid_m) == 0) next

    # ArbetstimmarGrund for this month
    hit_arb   <- arb[!is.na(arb$date) & arb$date == fd, , drop = FALSE]
    arb_hours <- if (nrow(hit_arb) > 0) {
      v <- suppressWarnings(as.numeric(hit_arb$arbetstimmar[1]))
      if (is.na(v)) 0 else v
    } else 0
    arb_hours_list[[sprintf("%04d%02d", y, m)]] <- arb_hours

    # Resolve timpris per Tidrapportering row using history
    tid_m$timpris_used <- vapply(seq_len(nrow(tid_m)), function(j) {
      resolve_timpris_for_month(
        timpris_history  = th_hist,
        uppgift_df       = uppgift,
        consultant_id    = tid_m$consultant_id[j],
        uppdrag_id       = tid_m$uppdrag_id[j],
        uppgift_id       = tid_m$uppgift_id[j],
        target_first_day = fd
      )
    }, numeric(1))

    tid_m$total_summa <- tid_m$timmar * tid_m$timpris_used

    # Attach consultant names, customer names and uppdrag names
    tid_m_named <- merge(tid_m,     kons_names,  by = "consultant_id", all.x = TRUE)
    tid_m_named <- merge(tid_m_named, kund_map,  by = "customer_id",   all.x = TRUE)
    tid_m_named <- merge(tid_m_named, uppdrag_map, by = "uppdrag_id",  all.x = TRUE)
    is_na_kund  <- is.na(tid_m_named$kundnamn)
    tid_m_named$kundnamn[is_na_kund]         <- tid_m_named$customer_id[is_na_kund]
    is_na_upp   <- is.na(tid_m_named$uppdrag_label)
    tid_m_named$uppdrag_label[is_na_upp]     <- tid_m_named$uppdrag_id[is_na_upp]

    # Detail: group by consultant + customer + uppdrag
    detail_m <- tid_m_named |>
      mutate(ar = y, manad = m) |>
      group_by(ar, manad, fornamn, efternamn, kundnamn, uppdrag_label) |>
      summarise(
        timmar      = sum(timmar,      na.rm = TRUE),
        total_summa = sum(total_summa, na.rm = TRUE),
        .groups = "drop"
      ) |>
      mutate(
        timpris_used    = ifelse(timmar > 0, total_summa / timmar, NA_real_),
        debiteringsgrad = ifelse(
          arb_hours > 0 & !is.na(timpris_used),
          round((total_summa / (arb_hours * timpris_used)) * 100, 2),
          NA_real_
        )
      ) |>
      select(ar, manad, fornamn, efternamn, kundnamn, uppdrag_label,
             timmar, timpris_used, total_summa, debiteringsgrad) |>
      arrange(fornamn, efternamn, kundnamn, uppdrag_label)

    # Resolve grundlön and bonus per unique consultant for this month
    cons_ids <- unique(tid_m$consultant_id)
    cons_resolved <- data.frame(
      consultant_id    = cons_ids,
      grundlon_used    = vapply(cons_ids, function(cid)
        resolve_grundlon_for_month(gl_hist, kons, cid, fd),       numeric(1)),
      bonus_grund_used = vapply(cons_ids, function(cid)
        resolve_bonus_for_month(bh_hist, kons, cid, fd), numeric(1)),
      stringsAsFactors = FALSE
    )
    # Attach names for joining to detail (which groups by fornamn/efternamn)
    cons_resolved <- merge(cons_resolved, kons_names, by = "consultant_id", all.x = TRUE)

    # Summary: group by consultant for this month
    summary_m <- detail_m |>
      group_by(ar, manad, fornamn, efternamn) |>
      summarise(
        total_debiterad = sum(total_summa,    na.rm = TRUE),
        debiteringsgrad = round(sum(debiteringsgrad, na.rm = TRUE), 2),
        .groups = "drop"
      ) |>
      left_join(
        cons_resolved |> select(fornamn, efternamn, grundlon_used, bonus_grund_used),
        by = c("fornamn", "efternamn")
      ) |>
      mutate(
        grundlon_used      = ifelse(is.na(grundlon_used),    0, grundlon_used),
        bonus_percent_used = ifelse(is.na(bonus_grund_used), 0, bonus_grund_used),
        bonus_rate         = bonus_percent_used / 100,
        bonus_belopp       = ifelse(total_debiterad > bonus_threshold,
                                    (total_debiterad - bonus_threshold) * bonus_rate,
                                    0),
        lon_total          = grundlon_used + bonus_belopp
      ) |>
      select(ar, manad, fornamn, efternamn, total_debiterad,
             grundlon_used, bonus_percent_used, bonus_belopp, lon_total, debiteringsgrad) |>
      arrange(fornamn, efternamn)

    monthly_detail[[i]]  <- detail_m
    monthly_summary[[i]] <- summary_m
  }

  monthly_detail  <- Filter(Negate(is.null), monthly_detail)
  monthly_summary <- Filter(Negate(is.null), monthly_summary)

  if (length(monthly_detail) == 0) return(make_empty())

  detail_all  <- bind_rows(monthly_detail)
  summary_all <- bind_rows(monthly_summary)

  # Interval totals: one row per consultant, summed across all months
  totals <- summary_all |>
    group_by(fornamn, efternamn) |>
    summarise(
      total_debiterad    = sum(total_debiterad,  na.rm = TRUE),
      grundlon_total     = sum(grundlon_used,    na.rm = TRUE),
      bonus_belopp_total = sum(bonus_belopp,     na.rm = TRUE),
      lon_total          = sum(lon_total,        na.rm = TRUE),
      debiteringsgrad    = round(sum(debiteringsgrad, na.rm = TRUE), 2),
      .groups = "drop"
    ) |>
    arrange(fornamn, efternamn)

  list(detail = detail_all, summary = summary_all, totals = totals,
       arb_hours_by_month = arb_hours_list)
}

# Wraps the three output tables for write_xlsx.
interval_report_workbook_list <- function(report) {
  list(
    Detalj         = report$detail,
    Sammanfattning = report$summary,
    Totalt         = report$totals
  )
}

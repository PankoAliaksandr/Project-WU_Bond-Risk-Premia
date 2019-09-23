library(readxl)
library(dplyr)
library(xtable)

data_load <- function(){
  price_data <- read_excel(file.choose(), 
                           col_types = c("numeric",
                                         "numeric",
                                         "numeric",
                                         "numeric",
                                         "numeric",
                                         "numeric"))
  
  price_data[, 1] <- as.Date(as.character(price_data$`Quotation date of observation`), format="%Y%m%d")
  return(price_data)
}
get_log_price_data <- function(price_data){
  # The first column is date
  date <- price_data[, 1]
  log_price_data <- log(price_data[, -1] / 100)
  log_price_data <- cbind(date, log_price_data)
  return(log_price_data)
}
get_log_yields <- function(log_price_data, n){
  # The first column is date
  log_yields <- as.data.frame(log_price_data[, 1])
  
  for (i in 2:n){
    log_yields[, i] <- - (log_price_data[, i])/ (i - 1) 
  }  
  
  colnames(log_yields) <- c("date", "y1", "y2", "y3", "y4", "y5")
  
  return(log_yields)
  
}
get_log_forward_rates <- function(log_price_data, n){
  
  # First column is date
  log_forwards <- as.data.frame(log_price_data[, 1])
  
  # Drop date column
  l_log_price_data <- log_price_data[, -1]
  
  # log_forwards(without 1 year bond)
  for (i in 2:(n-1)){
    log_forwards[, i] <- l_log_price_data[, (i - 1)] - l_log_price_data[, i]  
  }
  
  colnames(log_forwards) <- c("date", "f2","f3", "f4", "f5")
  
  return(log_forwards)
}
get_log_holding_period_returns <- function(log_price_data, r){
  
  # Save date column
  log_returns <- as.data.frame(log_price_data[, 1])
  
  # Next year starts in 12 months
  for (j in 13:r) {
    log_returns[j, 2] <- log_price_data[j, 2] - log_price_data[(j-12), 3] 
    log_returns[j, 3] <- log_price_data[j, 3] - log_price_data[(j-12), 4]
    log_returns[j, 4] <- log_price_data[j, 4] - log_price_data[(j-12), 5]
    log_returns[j, 5] <- log_price_data[j, 5] - log_price_data[(j-12), 6]
  }
  
  colnames(log_returns) <- c("date", "r2", "r3", "r4", "r5")
  
  return(log_returns)
}
get_excess_log_returns <- function(log_returns, log_yields, r){
  
  excess_log_ret <- as.data.frame(log_returns[, 1])
  
  for (j in 13:r) {
    excess_log_ret[j, 2] <- log_returns[j, 2] - log_yields[(j - 12), 2]
    excess_log_ret[j, 3] <- log_returns[j, 3] - log_yields[(j - 12), 2]
    excess_log_ret[j, 4] <- log_returns[j, 4] - log_yields[(j - 12), 2]
    excess_log_ret[j, 5] <- log_returns[j, 5] - log_yields[(j - 12), 2]
  }
  
  # average excess return across maturity
  summ = 0
  counter = 0
  for(i in 2:5){
    counter = counter + 1
    summ = summ + excess_log_ret[, i]
  }
  
  excess_log_ret[, 6] <- summ / counter
  
  colnames(excess_log_ret) <- c("date", "rx2", "rx3", "rx4", "rx5", "rxmean")
  
  # Delete rows with NA (first 12 empty rows)
  excess_log_ret <- excess_log_ret[rowSums(is.na(excess_log_ret)) == 0, ]
  
  return(excess_log_ret)
}
execute_unrestr_regressions <- function(excess_log_ret, log_yields, log_forwards, n_periods){
  
  coeff_u1 <- summary(lm(excess_log_ret[1:n_periods, 2] ~  log_yields[1:n_periods, 2] +
                                                     log_forwards[1:n_periods, 2] + 
                                                     log_forwards[1:n_periods, 3] +
                                                     log_forwards[1:n_periods, 4] +
                                                     log_forwards[1:n_periods, 5]))$coefficients
  
  coeff_u2 <- summary(lm(excess_log_ret[1:n_periods, 3] ~  log_yields[1:n_periods, 2] +
                                                     log_forwards[1:n_periods, 2] + 
                                                     log_forwards[1:n_periods, 3] +
                                                     log_forwards[1:n_periods, 4] +
                                                     log_forwards[1:n_periods, 5]))$coefficients
    
  coeff_u3 <- summary(lm(excess_log_ret[1:n_periods, 4] ~  log_yields[1:n_periods, 2] +
                                                     log_forwards[1:n_periods, 2] + 
                                                     log_forwards[1:n_periods, 3] +
                                                     log_forwards[1:n_periods, 4] +
                                                     log_forwards[1:n_periods, 5]))$coefficients
  
  coeff_u4 <- summary(lm(excess_log_ret[1:n_periods, 5] ~  log_yields[1:n_periods, 2] +
                                                     log_forwards[1:n_periods, 2] + 
                                                     log_forwards[1:n_periods, 3] +
                                                     log_forwards[1:n_periods, 4] +
                                                     log_forwards[1:n_periods, 5]))$coefficients
  if(n_periods == 480){
    plot_title <- "Unrestricted case for period until 2003"
  }else{
    plot_title <- "Unrestricted case for whole period"
  }
    
  
  plot(
    coeff_u1[2:6, 1],
    ylim = c(-4, 4.01),
    type = "b",
    lty = 2,
    lwd = 3,
    main = plot_title,
    xlab = "Gammas",
    ylab = ""
  )
  
  points(coeff_u2[2:6, 1],
         type = "b",
         pch = 25,
         col = "green")
  
  points(coeff_u3[2:6, 1],
         type = "b",
         pch = 19,
         col = "blue")
  
  points(coeff_u4[2:6, 1],
         type = "b",
         pch = 15,
         col = "red")
  
  legend(
    "topleft",
    c("2", "3", "4", "5"),
    lwd = 1,
    bty = "white",
    col = c("black", "green", "blue", "red"),
    cex = 0.8,
    y.intersp = 0.8,
    inset = 0.02,
    pch = c(19, 25, 19, 15)
  )
  
  betas <- matrix(ncol = 4, nrow = 6)
  betas[, 1] <- coeff_u1[, 1]
  betas[, 2] <- coeff_u2[, 1]
  betas[, 3] <- coeff_u3[, 1]
  betas[, 4] <- coeff_u4[, 1]
  
  xtable(betas, digits = 4)
  
  pvalues <- matrix(ncol = 4, nrow = 6)
  pvalues[, 1] <- coeff_u1[, 4]
  pvalues[, 2] <- coeff_u2[, 4]
  pvalues[, 3] <- coeff_u3[, 4]
  pvalues[, 4] <- coeff_u4[, 4]
  
  xtable(pvalues, digits = 4)
  
  return_list <- list(betas = betas,
                      pvalues = pvalues)
  
  return(return_list)
}
execute_restr_regressions <- function(excess_log_ret, log_yields, log_forwards, n_periods){
  
  # 1st step
  coeff_r1 <- summary(lm(excess_log_ret[1:n_periods, 6] ~  log_yields[1:n_periods, 2] + 
                           log_forwards[1:n_periods, 2] + 
                           log_forwards[1:n_periods, 3] + 
                           log_forwards[1:n_periods, 4] +
                           log_forwards[1:n_periods, 5]))$coefficients
  
  # 2nd step
  # log_forwards matrix
  forw <- matrix(nrow = 6, ncol = 612)
  forw[1, ] <- rep(1, 612)
  forw[2, ] <- log_yields[, 2]
  forw[3, ] <- log_forwards[, 2]
  forw[4, ] <- log_forwards[, 3]
  forw[5, ] <- log_forwards[, 4]
  forw[6, ] <- log_forwards[, 5]
  
  # gammas times f
  gammas <- as.matrix(coeff_r1[, 1])
  p_values_gammas <- as.matrix(coeff_r1[, 4])
  indepent_var <- t(gammas) %*% forw
  
  # regressions
  coeff_r21 <- summary(lm(excess_log_ret[1:n_periods, 2] ~ t(indepent_var)[1:n_periods] - 1))$coefficients
  coeff_r22 <- summary(lm(excess_log_ret[1:n_periods, 3] ~ t(indepent_var)[1:n_periods] - 1))$coefficients
  
  coeff_r23 <- summary(lm(excess_log_ret[1:n_periods, 4] ~ t(indepent_var)[1:n_periods] - 1))$coefficients
  coeff_r24 <- summary(lm(excess_log_ret[1:n_periods, 5] ~ t(indepent_var)[1:n_periods] - 1))$coefficients
  
  # multiplying coefficients( from 1st step) on coeff from 2nd step
  
  regfin <- matrix(nrow = 5, ncol = 4)
  for (i in 1:5) {
    regfin[i, 1] <- coeff_r21[, 1] * coeff_r1[i + 1, 1]
    regfin[i, 2] <- coeff_r22[, 1] * coeff_r1[i + 1, 1]
    regfin[i, 3] <- coeff_r23[, 1] * coeff_r1[i + 1, 1]
    regfin[i, 4] <- coeff_r24[, 1] * coeff_r1[i + 1, 1]
  }
  
  betas <- c(rep(0, times = 4))
  betas[1] <- coeff_r21[, 1]
  betas[2] <- coeff_r22[, 1]
  betas[3] <- coeff_r23[, 1]
  betas[4] <- coeff_r24[, 1]
  
  xtable(as.matrix(betas), digits = 4)
  
  p_values_betas <- c(rep(0, times = 4))
  p_values_betas[1] <- coeff_r21[, 4]
  p_values_betas[2] <- coeff_r22[, 4]
  p_values_betas[3] <- coeff_r23[, 4]
  p_values_betas[4] <- coeff_r24[, 4]
  
  xtable(as.matrix(p_values_betas), digits = 4)
  
  return_list <- list(gammas = gammas,
                      p_values_gammas = p_values_gammas,
                      betas = betas,
                      p_values_betas = p_values_betas,
                      regfin = regfin)
  
  # graph restricted
  if(n_periods == 480){
    plot_title <- "Restricted case for period until 2003"
  }else{
    plot_title <- "Restricted case for whole period"
  }
  plot(
    regfin[, 1],
    ylim = c(-4, 4.01),
    type = "b",
    lty = 2,
    lwd = 3,
    main = plot_title,
    xlab = "Gammas",
    ylab = ""
  )
  
  points(regfin[, 2],
         type = "b",
         pch = 25,
         col = "green")
  
  points(regfin[, 3],
         type = "b",
         pch = 19,
         col = "blue")
  
  points(regfin[, 4],
         type = "b",
         pch = 15,
         col = "red")
  
  legend(
    "topleft",
    c("2", "3", "4", "5"),
    lwd = 1,
    bty = "white",
    col = c("black", "green", "blue", "red"),
    cex = 0.8,
    y.intersp = 0.8,
    inset = 0.02,
    pch = c(19, 25, 19, 15)
  )
  
  return(return_list)
}

price_data <- data_load()

n <- ncol(price_data)
r <- nrow(price_data)

log_price_data <- get_log_price_data(price_data)
log_yields <- get_log_yields(log_price_data, n)
log_forwards <- get_log_forward_rates(log_price_data, n)
log_returns <- get_log_holding_period_returns(log_price_data, r)
excess_log_ret <- get_excess_log_returns(log_returns, log_yields, r)

n_periods_whole <- nrow(excess_log_ret)
partial_data <- subset(price_data, price_data[, 1] <= "2003/12/31")
n_periods_part <- nrow(partial_data)


unrestr_coeff_part <- execute_unrestr_regressions(excess_log_ret,
                                                   log_yields,
                                                   log_forwards,
                                                   n_periods = n_periods_part)

restr_coeff_part <- execute_restr_regressions(excess_log_ret,
                                               log_yields,
                                               log_forwards,
                                               n_periods = n_periods_part)

unrestr_coeff_whole <- execute_unrestr_regressions(excess_log_ret,
                                                     log_yields,
                                                     log_forwards,
                                                     n_periods = n_periods_whole)

restr_coeff_whole <- execute_restr_regressions(excess_log_ret,
                                                 log_yields,
                                                 log_forwards,
                                                 n_periods = n_periods_whole)

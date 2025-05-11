# ------------------ Load Libraries ------------------

library(openxlsx)
library(tidyverse)
library(mice)
library(naniar)
library(MissMech)
library(GGally)
library(caret)
library(ggcorrplot)
library(car)
library(corrplot)
library(skimr)
library(sandwich)
library(lmtest)
library(randomForest)

# ------------------ Load and Prepare Data ------------------

df <- read.xlsx("company_esg_financial_dataset.xlsx", sheet = 1, startRow = 2)

colnames(df) <- c("CompanyID", "CompanyName", "Industry", "Region", "Year", "Revenue", "ProfitMargin", "MarketCap", "GrowthRate", "ESG_Overall", "ESG_Environmental", "ESG_Social", "ESG_Governance", "CarbonEmissions", "WaterUsage", "EnergyConsumption")

df <- df %>% mutate(across(c(Revenue:EnergyConsumption), as.numeric))

# ------------------ Introduce Missingness ------------------

set.seed(123)
total_vals <- prod(dim(df)[1], 10)
n_missing <- floor(0.10 * total_vals)
missing_idx <- data.frame(row = sample(1:nrow(df), n_missing, replace = TRUE), col = sample(names(df)[6:15], n_missing, replace = TRUE))
for (i in 1:n_missing) {
  df[missing_idx$row[i], missing_idx$col[i]] <- NA
}

# ------------------ Save Dataset ------------------

write.csv(df, "company_esg_with_missing.csv", row.names = FALSE)

# ------------------ Visualize Missingness ------------------

vis_miss(df)
gg_miss_upset(df)
gg_miss_var(df)
gg_miss_fct(df, fct = Region)

# ------------------ Explore Missingness Mechanism ------------------

df_test <- df %>% select(where(is.numeric))
cor_matrix <- cor(df_test, use = "pairwise.complete.obs")
which(abs(cor_matrix) > 0.9999 & abs(cor_matrix) < 1, arr.ind = TRUE)
df_test <- df_test %>% select(-EnergyConsumption)
TestMCARNormality(df_test)

# ------------------ Imputation ------------------

vars_for_imputation <- c("ProfitMargin", "GrowthRate", "ESG_Overall", "ESG_Environmental", "ESG_Social", "ESG_Governance")
df_mice <- df %>% select(all_of(vars_for_imputation))
pred_mat <- make.predictorMatrix(df_mice)
diag(pred_mat) <- 0
imp <- mice(df_mice, method = "pmm", predictorMatrix = pred_mat, m = 5, maxit = 5, seed = 42)
df[vars_for_imputation] <- complete(imp, 1)
df <- df %>% drop_na()

# ------------------ Feature Engineering ------------------

df <- df %>% mutate(
  Profit = Revenue * ProfitMargin / 100,
  ESG_Spread = ESG_Governance - ESG_Environmental,
  EmissionsPerRevenue = CarbonEmissions / Revenue,
  EmissionsPerProfit = CarbonEmissions / (Profit + 1),
  ESG_High = ifelse(ESG_Overall >= median(ESG_Overall), 1, 0)
)

# ------------------ Exploratory Data Analysis ------------------

skim(df)
df %>%
  pivot_longer(cols = where(is.numeric)) %>%
  ggplot(aes(x = value)) +
  geom_histogram(bins = 30, fill = "skyblue") +
  facet_wrap(~name, scales = "free") +
  theme_minimal()
GGally::ggpairs(df %>% select(where(is.numeric)))

# ------------------ Statistical Tests ------------------

anova_result <- aov(ESG_Overall ~ Industry, data = df)
summary(anova_result)
tukey_result <- TukeyHSD(anova_result)
plot(tukey_result, las = 1, col = "steelblue")

# ------------------ Modeling ------------------

df <- df %>% mutate(log_Revenue = log(Revenue + 1), log_CarbonEmissions = log(CarbonEmissions + 1))
model_final <- lm(ESG_Overall ~ log_Revenue + log_CarbonEmissions + Industry + Region, data = df)
summary(model_final)
vif(model_final)
bptest(model_final)
coeftest(model_final, vcov = vcovHC(model_final, type = "HC1"))

# ------------------ Train-Test Evaluation ------------------

set.seed(42)
split <- createDataPartition(df$ESG_Overall, p = 0.8, list = FALSE)
train <- df[split, ]
test <- df[-split, ]
train_pred <- predict(model_final, newdata = train)
test_pred <- predict(model_final, newdata = test)
train_metrics <- postResample(train_pred, train$ESG_Overall)
test_metrics <- postResample(test_pred, test$ESG_Overall)

# ------------------ Visualization ------------------

df_pred <- data.frame(Actual = test$ESG_Overall, Predicted = test_pred)
ggplot(df_pred, aes(x = Actual, y = Predicted)) +
  geom_point(alpha = 0.4, color = "steelblue") +
  geom_abline(slope = 1, intercept = 0, color = "darkred", linetype = "dashed") +
  theme_minimal() +
  labs(title = "Predicted vs Actual ESG Score", x = "Actual", y = "Predicted") +
  coord_equal()

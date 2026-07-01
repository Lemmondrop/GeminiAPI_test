param(
  [string]$Version = "v$(Get-Date -Format 'yyyy_MM_dd_HHmm')"
)

$ErrorActionPreference = "Stop"

$Base = Join-Path "test_data" $Version
$Raw = Join-Path $Base "01_raw"
$Phase = Join-Path $Base "02_phase_outputs"
$Training = Join-Path $Base "03_training"
$Models = Join-Path $Base "04_models"
$XmlDir = Join-Path $Base "xml_peer_group"

New-Item -ItemType Directory -Force -Path $Raw, $Phase, $Training, $Models, $XmlDir | Out-Null

$DiscountTables = "C:\Users\Researcher\Desktop\Project V\OCR Sample\data\밸류추정 인수인의 의견\dicount_rate_table"

$DiscountTableCsv = Join-Path $Raw "discount_rate_table_pages.csv"
$ExpandedRates = Join-Path $Raw "peer_group_discount_rates_expanded.csv"
$WithCorp = Join-Path $Phase "peer_group_expanded_with_corp_code.csv"
$WithRcpno = Join-Path $Phase "peer_group_expanded_with_rcpno.csv"
$FeaturesRaw = Join-Path $Phase "peer_group_expanded_features_raw.csv"
$FeaturesEnriched = Join-Path $Phase "peer_group_expanded_features_enriched.csv"
$TrainingCsv = Join-Path $Training "training_features_expanded.csv"
$SegmentedCsv = Join-Path $Training "training_features_expanded_segmented.csv"
$GeneralCsv = Join-Path $Training "training_features_expanded_general.csv"
$GeneralExcludedCsv = Join-Path $Training "training_features_expanded_general_excluded.csv"

python extract_halla_discount_table.py `
  --input-dir $DiscountTables `
  --glob "*_discount_rate_table.pdf" `
  --output $DiscountTableCsv `
  --method auto `
  --dedupe-with-rates

python merge_discount_rate_sources.py `
  --base peer_group_discount_rates.csv `
  --add $DiscountTableCsv `
  --output $ExpandedRates

python phase1_map_corp_codes.py `
  --input $ExpandedRates `
  --output $WithCorp

python phase2_find_rcpno.py `
  --input $WithCorp `
  --output $WithRcpno `
  --rate-limit-sec 0.3

python phase3_extract_features.py `
  --input $WithRcpno `
  --output $FeaturesRaw `
  --xml-dir $XmlDir `
  --rate-limit-sec 0.5

python phase3_5_dart_financials.py `
  --input $FeaturesRaw `
  --output $FeaturesEnriched `
  --only-missing

python build_training_features.py `
  --inputs $FeaturesEnriched ipo_features_raw.csv `
  --output $TrainingCsv `
  --stock-code-maps $WithCorp ipo_with_corp_code_v2.csv ipo_with_corp_code.csv

python build_crawling_segments.py `
  --input $TrainingCsv `
  --output $SegmentedCsv `
  --cache-dir crawling_cache

python build_general_training_features.py `
  --input $SegmentedCsv `
  --output $GeneralCsv `
  --excluded-output $GeneralExcludedCsv

$env:TRAINING_CSV = $GeneralCsv
$env:MODEL_OUT = Join-Path $Models "model_general.pkl"
$env:IMPORTANCE_OUT = Join-Path $Models "feature_importance_general.csv"
$env:PREDICTIONS_OUT = Join-Path $Models "predictions_cv_general.csv"
$env:METRICS_OUT = Join-Path $Models "model_metrics_general.csv"
python phase4_train_model_general.py

$env:MODEL_OUT = Join-Path $Models "model_input_limited.pkl"
$env:IMPORTANCE_OUT = Join-Path $Models "feature_importance_input_limited.csv"
$env:PREDICTIONS_OUT = Join-Path $Models "predictions_cv_input_limited.csv"
$env:METRICS_OUT = Join-Path $Models "model_metrics_input_limited.csv"
python phase4_train_model_input_limited.py

Remove-Item Env:TRAINING_CSV -ErrorAction SilentlyContinue
Remove-Item Env:MODEL_OUT -ErrorAction SilentlyContinue
Remove-Item Env:IMPORTANCE_OUT -ErrorAction SilentlyContinue
Remove-Item Env:PREDICTIONS_OUT -ErrorAction SilentlyContinue
Remove-Item Env:METRICS_OUT -ErrorAction SilentlyContinue

python snapshot_dataset_version.py --version $Version

Write-Host ""
Write-Host "Versioned pipeline complete: $Base"

// List of locales known to use comma decimal separator.
const localesUseCommaSep = new Set([
  'ar_EG', // Egypt
  'az_AZ', // Azerbajian
  'be_BY', // Belarus
  'bg_BG', // Bulgaria
  'cs_CZ', // Czechia
  'da_DK', // Denmark
  'de_DE', // Germany
  'el_GR', // Greece
  'es_AR', // Argentina
  'es_BO', // Bolivia
  'es_CL', // Chile
  'es_CO', // Colombia
  'es_EC', // Ecuador
  'es_ES', // Spain
  'es_PY', // Paraguay
  'es_UY', // Uruguay
  'es_VE', // Venezuela
  'fi_FI', // Finland
  'fr_CA', // Canada (French)
  'fr_FR', // France
  'hr_HR', // Croatia
  'hu_HU', // Hungary
  'hy_AM', // Armenia
  'id_ID', // Indonesia
  'it_IT', // Italy
  'kk_KZ', // Kazakhstan
  'ka_GE', // Georgia
  'lt_LT', // Lithuania
  'lv_LV', // Latvia
  'nb_NO', 'nn_NO', // Norway
  'nl_NL', // Netherlands
  'pl_PL', // Poland
  'pt_BR', // Brazil
  'pt_PT', // Portugal
  'ro_RO', // Romania
  'sk_SK', // Slovakia
  'sl_SI', // Slovenia
  'sr_RS', // Serbia
  'sv_SE', // Sweden
  'tr_TR', // Turkey
  'ru_RU', // Russia
  'uk_UA', // Ukraine
  'vi_VN', // Vietnam
]);

function useCommaDecimalSep() {
  const locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  return localesUseCommaSep.has(locale);
};

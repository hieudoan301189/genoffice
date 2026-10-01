import { aiStrings } from './strings-ai'
import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { dvhStrings } from './strings-dvh'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...aiStrings.zh, ...dvhStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...aiStrings.en, ...dvhStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...aiStrings.ja, ...dvhStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...aiStrings.ko, ...dvhStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...aiStrings.fr, ...dvhStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...aiStrings.de, ...dvhStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...aiStrings.es, ...dvhStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...aiStrings.th, ...dvhStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...aiStrings.id, ...dvhStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...aiStrings.ru, ...dvhStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...aiStrings.ar, ...dvhStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...aiStrings.pt, ...dvhStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...aiStrings.it, ...dvhStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...aiStrings.pl, ...dvhStrings.pl },
  cs: { ...appStrings.cs, ...dialogStrings.cs, ...aiStrings.cs, ...dvhStrings.cs },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...aiStrings.nl, ...dvhStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...aiStrings.ms, ...dvhStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...aiStrings.he, ...dvhStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...aiStrings.hi, ...dvhStrings.hi },
  'zh-TW': { ...appStrings['zh-TW'], ...dialogStrings['zh-TW'], ...aiStrings['zh-TW'], ...dvhStrings['zh-TW'] },
}

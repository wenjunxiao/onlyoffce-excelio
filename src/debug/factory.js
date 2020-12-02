/* eslint no-console: off */
export default function debugFactory (name) {
  if (typeof window !== 'undefined' && window.localStorage) {
    const reg = new RegExp(`(?:^|,)\\s*${name}(?:,\\s*|$)`, 'i');
    if (reg.test(window.localStorage.getItem('debug'))) {
      return console.log.bind(console);
    }
  }
  return () => { }
}
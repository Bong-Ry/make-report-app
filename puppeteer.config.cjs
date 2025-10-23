const { join } = require('path');

/**
 * @type {import("puppeteer").Configuration}
 */
module.exports = {
  // PuppeteerがChromiumをダウンロードする場所を、
  // プロジェクト内の .cache/puppeteer フォルダに変更する
  cacheDirectory: join(__dirname, '.cache', 'puppeteer'),
};

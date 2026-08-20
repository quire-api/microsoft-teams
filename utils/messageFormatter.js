// Copyright (C) 2026 Potix Corporation. All Rights Reserved
// History: 2026/08/20
// Author: jimmy<jimmyshiau@potix.com>

const { AllHtmlEntities } = require('html-entities');

const entities = new AllHtmlEntities();

/**
 * Converts the HTML markup in a Quire notification message to Teams
 * markdown, so bot notifications never expose raw tags
 * (MS certification policy 1140.4.3.1, boeneo#25590).
 *
 * Tag/status chips (`<code class="tag iconc-NN">Name</code>`) render as
 * the bare name — `iconc-NN` palette indices are Quire internals and must
 * not leak into Teams. Closing tags carrying attributes (as seen in the
 * certification screenshots) are stripped the same as well-formed ones.
 */
function htmlToTeamsMarkdown(html) {
  if (!html) return '';
  let text = String(html);

  // tag / status / priority chips: keep the name, drop the chip markup
  text = text.replace(/<\/?code\b[^>]*>/gi, '');

  // links
  text = text.replace(
    /<a\b[^>]*href=(["'])([\s\S]*?)\1[^>]*>([\s\S]*?)<\/a\b[^>]*>/gi,
    '[$3]($2)');

  // bold / italic
  text = text.replace(/<\/?(?:b|strong)\b[^>]*>/gi, '**');
  text = text.replace(/<\/?(?:i|em)\b[^>]*>/gi, '_');

  // line breaks
  text = text.replace(/<br\b[^>]*\/?>/gi, '\n');

  // anything else Quire may emit: strip, keep the inner text
  text = text.replace(/<\/?[a-zA-Z][^>]*>/g, '');

  return entities.decode(text).trim();
}

module.exports = {
  htmlToTeamsMarkdown: htmlToTeamsMarkdown
}

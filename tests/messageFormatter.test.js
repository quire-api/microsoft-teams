// Copyright (C) 2026 Potix Corporation. All Rights Reserved
// History: 2026/08/20
// Author: jimmy<jimmyshiau@potix.com>

const test = require('node:test');
const assert = require('node:assert');
const { htmlToTeamsMarkdown } = require('../utils/messageFormatter');

test('plain text passes through unchanged', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown('Qure Dev commented mentioned comment on Task 1'),
    'Qure Dev commented mentioned comment on Task 1');
});

test('empty and missing messages', () => {
  assert.strictEqual(htmlToTeamsMarkdown(''), '');
  assert.strictEqual(htmlToTeamsMarkdown(null), '');
  assert.strictEqual(htmlToTeamsMarkdown(undefined), '');
});

test('status chip renders as bare name (SetState)', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown(
      'Qure Dev set the status of Subtask 1 to <code class="tag iconc-43">Completed</code>'),
    'Qure Dev set the status of Subtask 1 to Completed');
});

test('malformed closing tag carrying attributes is stripped too', () => {
  // shape observed in the certification screenshot (boeneo#25590)
  assert.strictEqual(
    htmlToTeamsMarkdown(
      'Qure Dev set the status of Subtask 1 to <code class="tag iconc-43">Completed</code class="tag iconc-43">'),
    'Qure Dev set the status of Subtask 1 to Completed');
});

test('tag chips render as bare names (SetTag)', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown(
      'Qure Dev added the tag <code class="tag iconc-22">New</code> to Subtask 2'),
    'Qure Dev added the tag New to Subtask 2');
});

test('links become markdown links', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown(
      'Qure Dev completed <a href="https://quire.io/w/my_project/123">Task 1</a>'),
    'Qure Dev completed [Task 1](https://quire.io/w/my_project/123)');
});

test('bold and italic become markdown', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown('<b>Qure Dev</b> edited <i>Task 1</i>'),
    '**Qure Dev** edited _Task 1_');
  assert.strictEqual(
    htmlToTeamsMarkdown('<strong>Qure Dev</strong> edited <em>Task 1</em>'),
    '**Qure Dev** edited _Task 1_');
});

test('br becomes newline', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown('line 1<br>line 2<br/>line 3'),
    'line 1\nline 2\nline 3');
});

test('unknown tags are stripped, inner text kept', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown('added <span class="x">a note</span> to <u>Task 1</u>'),
    'added a note to Task 1');
});

test('html entities are decoded', () => {
  assert.strictEqual(
    htmlToTeamsMarkdown('renamed <b>R&amp;D &lt;draft&gt;</b>'),
    'renamed **R&D <draft>**');
});

test('no iconc palette index ever reaches the output', () => {
  const samples = [
    'to <code class="tag iconc-43">Completed</code>',
    'to <code class="tag iconc-43">Completed</code class="tag iconc-43">',
    'the tag <code class="tag iconc-24">Urgent</code> to Subtask 2',
  ];
  for (const s of samples)
    assert.ok(!/iconc-\d+/.test(htmlToTeamsMarkdown(s)), s);
});

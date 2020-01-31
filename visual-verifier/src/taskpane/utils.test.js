const paraLimiter = require('./utils');

test('adds 1 to equal 3', () => {
  expect(paraLimiter('1')).toBe(3);
});
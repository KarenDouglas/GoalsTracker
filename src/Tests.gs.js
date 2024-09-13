// Optional for easier use.
let QUnit = QUnitGS2.QUnit;

function doGet() {
  QUnitGS2.init(); // Initializes the library.

  QUnit.module("Math Functions Tests");



QUnit.test('square()', assert => {
  assert.equal(square(2), 4);
  assert.equal(square(3), 9);
  Logger.log(  assert.equal(square(2), 4))
});


  QUnit.start(); // Starts running tests.
  return QUnitGS2.getHtml();
}
function square (x) {
  return x * x;
}

function getResultsFromServer() {
  return QUnitGS2.getResultsFromServer();
}


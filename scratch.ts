const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

const getColumnFromIndex = (index: number) => {
  if (index >= LETTERS.length * LETTERS.length + LETTERS.length) {
    throw new Error(`Column index conversion exceeds maximum length: ${index}`);
  }

  let column = '';

  const magnitude = Math.floor(index / LETTERS.length);

  if (magnitude) {
    column += LETTERS[magnitude - 1];
  }

  column += LETTERS[index % LETTERS.length];

  return column;
};

const getIndexFromColumn = (column: string) => {
  const placeValues = column.split('').reverse();

  if (placeValues.length > 2) {
    throw new Error(
      `Index to column conversion exceeds maximum input: ${column}`
    );
  }

  let index = 0;

  index += LETTERS.indexOf(placeValues[0]);

  if (placeValues[1]) {
    index += (LETTERS.indexOf(placeValues[1]) + 1) * LETTERS.length;
  }

  return index;
};

console.log('A', getIndexFromColumn('A'));
console.log('B', getIndexFromColumn('B'));
console.log('Z', getIndexFromColumn('Z'));
console.log('AA', getIndexFromColumn('AA'));
console.log('AB', getIndexFromColumn('AB'));
console.log('AZ', getIndexFromColumn('AZ'));
console.log('BA', getIndexFromColumn('BA'));

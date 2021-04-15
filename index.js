// Make values inaccessible from outside
(function () {
  const moveBorder = (e) => {
    // Define selection
    if (e.target.getAttribute('data-status') === 'false') {
      const ref = document.getElementsByClassName('input-cell');

      // Find element with 'true' value of data-status
      for (let i = 0; i < ref.length; i++) {
        if (ref[i].getAttribute('data-status') === 'true') {
          ref[i].setAttribute('data-status', 'false');
          ref[i].style = 'border: none;';

          break;
        }
      }

      // change style for target
      e.target.setAttribute('data-status', 'true');

      e.target.style = 'border: 2px solid #000000; border-radius: 5px;';
    }
  };

  // Output to field
  const outputCurrentCell = () => {
    // output to formula field
    if (document.activeElement.value) {
      document.getElementById('input-formula').value =
        document.activeElement.value;
    } else {
      document.getElementById('input-formula').value = '';
    }

    // output current cell
    document.getElementById('sel-current').value = document.activeElement.id;
  };

  // One cell input
  document.getElementById('input-formula').addEventListener(
    'keyup',
    () => {
      const ref = document.getElementsByClassName('input-cell');
      const cellFormula = document.getElementById('input-formula').value;

      // Find element with 'true' value of data-status
      for (let i = 0; i < ref.length; i++) {
        if (ref[i].getAttribute('data-status') === 'true') {
          ref[i].value = cellFormula;

          break;
        }
      }
    },
    false
  );

  // Check Formula
  const getValue = (e) => {
    const id = e.target.id;
    const input = e.target.value.replace(/\s*/g, '');

    // if "Enter" is pressed
    if (window.event.keyCode === 13) {
      const formulaPattern = /(^=SUM\(([A-J]\d{1,2}[:/*+-])(([A-J]\d{1,2}\))$))|(^=([A-J]\d{1,2}[:/*+-])([A-J]\d{1,2}))/gi;

      // Check if its formula input
      if (formulaPattern.test(input)) {
        // Sanitize and split input
        const splitPattern = /[=SUM=\(\)]/gi;
        const splittedArr = input
          .split(splitPattern)
          .filter(Boolean)[0]
          .split('');

        // split each portion of cells
        for (let i = 0; i < splittedArr.length; i++) {
          if (/[:/*+-]/.test(splittedArr[i])) {
            // call cellsLogic to do math and return function
            return cellsLogic({
              firstCell: splittedArr.slice(0, i).join(''),
              secondCell: splittedArr.slice(i + 1, splittedArr.length).join(''),
              sign: splittedArr[i],
              id: 'input-formula',
            });
          }
        }
      }
    }
  };

  const spawnTable = () => {
    // Declare main variables
    const ROWS = 20;
    const COLUMNS = 10;
    const letters = ['#', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
    const refToExcelWorksheet = document.getElementById('excel-worksheet');

    // Func to create DOM element
    const domElement = (element) => document.createElement(element);

    // For return func: map list, history
    const inputList = new Map();

    // Make table
    const table = domElement('table');
    const thead = domElement('thead');
    const tbody = domElement('tbody');

    table.appendChild(thead);
    table.appendChild(tbody);

    for (let b = 0; b <= ROWS; b++) {
      const row = domElement('tr');

      for (let a = 0; a <= COLUMNS; a++) {
        let cell = undefined;

        if (b === 0) {
          cell = domElement('th');
          cell.textContent = letters[a];
        } else {
          cell = domElement('td');

          if (a === 0) {
            cell.textContent = b;
          } else {
            const input = domElement('input');

            // Display currently selected cell
            input.addEventListener('click', outputCurrentCell, false);
            input.addEventListener('click', moveBorder, false);

            // Check for func
            input.addEventListener('keypress', getValue, false);

            // Validate cell value for digits
            input.addEventListener(
              'keyup',
              (e) => {
                if (window.event.keyCode !== 13) {
                  outputCurrentCell();
                }

                const pattern = /^[0-9]+$/;

                if (e.target.value.match(pattern)) {
                  return cellsLogic({
                    firstCell: null,
                    secondCell: null,
                    sign: null,
                    id: e.target.id,
                  });
                }
              },
              false
            );

            // Generate ID with letter and digit: ex. A1
            input.id = `${letters[a]}${b}`;

            input.setAttribute('data-status', 'false');

            // Set mapid and reference to its cell
            inputList.set(input.id, `document.getElementById('${input.id}')`);

            input.className = 'input-cell clear-box';
            cell.appendChild(input);
          }
        }

        row.appendChild(cell);
      }

      if (b === 0) {
        thead.appendChild(row);
      } else {
        tbody.appendChild(row);
      }
    }

    refToExcelWorksheet.appendChild(table);

    // return mapped cells to outputInCell
    return inputList;
  };

  // Cells output
  const outputInCell = ((inputList = spawnTable()) => {
    // Encapsulate
    const toExecute = (command) => eval(command);
    const getCommand = (key) => inputList.get(key);
    const formula = document.getElementById('input-formula');

    return (arrCells, sign, refOutput, history) => {
      const executeRef = [];

      if (sign === ':') {
        sign = '+';
      }

      // Convert cells to executable references
      for (let i = 0; i < arrCells.length; i++) {
        executeRef.push(getCommand(arrCells[i]));
      }

      refOutput = toExecute(getCommand(refOutput));

      if (!arrCells.indexOf(refOutput.id)) {
        formula.value = 'Please select another field or change params!';
        history.pop();
      } else {
        const valArr = [];

        for (let i = 0; i < executeRef.length; i++) {
          const element = /^\d+$/.test(toExecute(executeRef[i]).value)
            ? parseFloat(toExecute(executeRef[i]).value)
            : 0;
          valArr.push(element);
        }

        const sum = valArr.reduce((sum, i) => toExecute(sum + sign + i));
        console.log(sum);

        for (let i = 0; i < executeRef.length; i++) {
          toExecute(executeRef[i]).value = valArr[i];
        }

        // Output for the cells
        refOutput.value = sum;
        formula.value = sum;
      }

      return history;
    };
  })();

  // Cells logic
  const cellsLogic = (() => {
    let history = [];
    const cells = document.getElementsByClassName('input-cell');

    return (inputProcess) => {
      // Check if clear button is pressed
      if (inputProcess === 'click') {
        return (history = []);
      }

      // Define statement values
      const arrCells = [];
      let { firstCell, secondCell, sign, id } = inputProcess;
      let refOutput = '';

      // Check if user inserted formula
      if (id === 'input-formula') {
        // Try to navigate cell for output
        for (let i = 0; i < cells.length; i++) {
          if (cells[i].getAttribute('data-status') === 'true') {
            refOutput = cells[i].id;
          }
        }

        if (sign === ':') {
          if (firstCell.charAt(0) === secondCell.charAt(0)) {
            const length =
              firstCell.charAt(1) >= secondCell.charAt(1)
                ? firstCell.charAt(1)
                : secondCell.charAt(1);
            for (let i = 0; i < length; i++) {
              arrCells.push(`${firstCell.charAt(0)}${i + 1}`);
            }
          } else if (
            firstCell.charAt(0) !== secondCell.charAt(0) &&
            firstCell.charAt(1) === secondCell.charAt(1)
          ) {
            const letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
            const firstIndex = letters.indexOf(firstCell.charAt(0));
            const secondIndex = letters.indexOf(secondCell.charAt(0));
            const length =
              firstIndex >= secondIndex
                ? firstIndex - secondIndex
                : secondIndex - firstIndex;

            for (let i = 0; i <= length; i++) {
              arrCells.push(`${letters[i]}${firstCell.charAt(1)}`);
            }
          }
        } else {
          arrCells.push(firstCell);
          arrCells.push(secondCell);
        }

        // Make history notice
        history.push({
          hCells: arrCells,
          hSign: sign,
          hRefOutput: refOutput,
        });

        // Call to display logics and renew history
        history = outputInCell(arrCells, sign, refOutput, history);
      } else {
        let refArr = '';
        let isIn = false;

        // Check if we have already inserted any data with these cells
        for (let i = 0; i < history.length; i++) {
          for (let b = 0; b < history[i].hCells.length; b++) {
            if (history[i].hCells[b] == id) {
              refArr = history[i];
              isIn = true;
              break;
            }
          }
        }

        // if its in, we call outputInCell, where we calculate and output
        if (isIn) {
          const { hCells, hSign, hRefOutput } = refArr;

          console.log(history);
          history = outputInCell(hCells, hSign, hRefOutput, history);
          console.log(history);
        }
      }
    };
  })();

  // Load some basic event listeners
  const loadBasics = (() => {
    // Check formula
    document
      .getElementById('input-formula')
      .addEventListener('keypress', getValue, false);

    // Clear all boxes
    document.getElementsByClassName('icon-btn')[0].addEventListener(
      'click',
      () => {
        // Clear history
        cellsLogic('click');

        // Clear cells and formula input
        document
          .querySelectorAll('.clear-box')
          .forEach((el) => (el.value = ''));
      },
      false
    );
  })();
})();

import Cell from "./Cell";
import SheetMemory from "./SheetMemory";
import { ErrorMessages } from "./GlobalDefinitions";

export class FormulaEvaluator {
  // Define a function called update that takes a string parameter and returns a number
  private _errorOccured: boolean = false;
  private _errorMessage: string = "";
  private _currentFormula: FormulaType = [];
  private _lastResult: number = 0;
  private _sheetMemory: SheetMemory;
  private _result: number = 0;

  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  /**
   * reset the evaluator
   */
  reset() {
    this._currentFormula = [];
    this._errorOccured = false;
    this._errorMessage = "";
    this._lastResult = 0;
    this._result = 0;
  }

  /**
   * evaluate the formula by recursive descent parser
   * @param formula
   * @returns the value of the expression in the tokenized formula
   *
   * If the formula is valid, returns the value of the expression in the formula
   * Otherwise sets the errorOccured flag to true and sets the errorMessage
   * The empty formula is considered as an emptyFormula error
   */

  evaluate(formula: FormulaType) {
    // reset the evaluator for the new formula
    this.reset();
    this._currentFormula = [...formula];

    // if the formula is empty set the errorOccured flag to true
    if (this._currentFormula.length === 0) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.emptyFormula;
      return;
    }

    // get the value of the expression in the formula
    // expression = term {("+" | "-") term}
    this._result = this.expression();

    if (this._errorOccured) {
      return;
    }

    // if there are still tokens in the formula, set the errorOccured flag to true and set the errorMessage
    if (this._currentFormula.length > 0) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
    }
  }

  /**
   * expression = term {("+" | "-") term}
   * term does not contain "+"/"-
   * @returns the value of the expression
   */
  private expression(): number {
    // extract the first term
    // term = factor {("*" | "/") factor}
    let result = this.term();
    // resolve "+"/"-" operators
    while (
      this._currentFormula.length > 0 &&
      (this._currentFormula[0] === "+" || this._currentFormula[0] === "-")
    ) {
      // extract the operator
      let operator = this._currentFormula.shift();
      // extract the next term
      let term = this.term();
      if (operator === "+") {
        result += term;
      } else {
        result -= term;
      }
    }
    return result;
  }

  /**
   * term = factor {("*" | "/") factor}
   * factor does not contain "*"/"/"
   * @returns the value of the term
   */
  private term(): number {
    // extract the first factor
    let result = this.factor();
    while (
      this._currentFormula.length > 0 &&
      (this._currentFormula[0] === "*" || this._currentFormula[0] === "/")
    ) {
      // extract the operator
      let operator = this._currentFormula.shift();
      // extract the next factor
      let factor = this.factor();
      if (operator === "*") {
        result *= factor;
      } else {
        if (factor === 0) {
          this._errorOccured = true;
          this._errorMessage = ErrorMessages.divideByZero;
          return Infinity;
        }
        result /= factor;
      }
    }
    return result;
  }

  /**
   * factor = number | (expression) | cellReference
   * anything else is an error
   * @returns the value of the factor
   */
  private factor(): number {
    let result = 0;
    // if the formula is empty, set as partial error
    if (this._currentFormula.length === 0) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.partial;
      return result;
    }

    // extract the first token from the formula
    let token = this._currentFormula.shift();

    // if the token is a number, set the value of the factor to the number
    if (this.isNumber(token)) {
      result = Number(token);
      // if the token is a "(" get the value of the expression
    } else if (token === "(") {
      // extract the whole expression in the paratheses
      result = this.expression();
      if (
        this._currentFormula.length === 0 ||
        this._currentFormula.shift() !== ")"
      ) {
        this._errorOccured = true;
        this._errorMessage = ErrorMessages.missingParentheses;
      }

      // if the token is a cell reference get the value of the cell
    } else if (this.isCellReference(token)) {
      [result, this._errorMessage] = this.getCellValue(token);

      // if the cell value is invalid, raise the error flag
      if (this._errorMessage !== "") {
        this._errorOccured = true;
      }

      // invalid token
    } else {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
    }
    return result;
  }

  public get error(): string {
    return this._errorMessage;
  }

  public get result(): number {
    return this._result;
  }

  /**
   *
   * @param token
   * @returns true if the toke can be parsed to a number
   */
  isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  /**
   *
   * @param token
   * @returns true if the token is a cell reference
   *
   */
  isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

  /**
   *
   * @param token
   * @returns [value, ""] if the cell formula is not empty and has no error
   * @returns [0, error] if the cell has an error
   * @returns [0, ErrorMessages.invalidCell] if the cell formula is empty
   *
   */
  getCellValue(token: TokenType): [number, string] {
    let cell = this._sheetMemory.getCellByLabel(token);
    let formula = cell.getFormula();
    let error = cell.getError();

    // if the cell has an error return 0
    if (error !== "" && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    // if the cell formula is empty return 0
    if (formula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }

    let value = cell.getValue();
    return [value, ""];
  }
}

export default FormulaEvaluator;

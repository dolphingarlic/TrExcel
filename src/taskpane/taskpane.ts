// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import { movePiece, rotatePiece } from "./actions";
import { BlockColorsMap, Directions, Height, Scores, Width } from "./constants";
import { getFreshSheet } from "./excel-utils";
import { BlockColors, Piece, randomPiece } from "./pieces";

/* global console, document, Excel, Office, clearInterval, setInterval */

//-------------------------------------------------------------------------
// game variables (initialized during reset)
//-------------------------------------------------------------------------

let blocks: (BlockColors | null)[][];
let playing: boolean = false;
let interval: any;

let score: number;
let currentPiece: Piece;
let nextPiece: Piece;
let pocketPiece: Piece;
let canPocket: boolean;

//------------------------------------------------
// Do the bit manipulation and iterate through each
// occupied block (x,y) for a given piece
//------------------------------------------------

function eachBlock(piece: Piece, fn: { (x: any, y: any): void; (arg0: any, arg1: any): void }) {
  let bitmask = piece.pieceType.blocks[piece.direction];
  for (let bit = 0x8000, block = 0; bit > 0; bit = bit >> 1, block++)
    if (bitmask & bit) fn(piece.x + Math.floor(block / 4), piece.y + (block % 4));
}

//-----------------------------------------------------
// Check if a piece can fit into a position in the grid
//-----------------------------------------------------

export function occupied(piece: Piece) {
  let result = false;
  eachBlock(piece, function(x: number, y: number) {
    if (x < 0 || x >= Width || y < 0 || y >= Height || blocks[x][y] != null) result = true;
  });
  return result;
}

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("loading-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    onkeydown = handleKeyPress;

    document.getElementById("start-game-button").onclick = resetGame;
    document.getElementById("pause-game-button").onclick = pauseGame;
    document.getElementById("resume-game-button").onclick = playGame;
  }
});

function cycleCurrentPiece() {
  if (!nextPiece) nextPiece = randomPiece();
  currentPiece = nextPiece;
  nextPiece = randomPiece();
}

async function resetGame() {
  // Get a fresh sheet and game loop
  await getFreshSheet();
  if (interval != null) clearInterval(interval);

  // Hide unneeded DOM elements
  document.getElementById("game-over-message").style.display = "none";
  document.getElementById("pause-game-button").style.display = "flex";
  document.getElementById("start-game-button").style.display = "none";

  blocks = new Array(Width);
  for (let i = 0; i < Width; i++) blocks[i] = new Array(Height).fill(null);

  score = 0;
  currentPiece = null;
  nextPiece = null;
  pocketPiece = null;
  canPocket = true;
  cycleCurrentPiece();
  playGame();
}

function pauseGame() {
  playing = false;
  clearInterval(interval);
  document.getElementById("pause-game-button").style.display = "none";
  document.getElementById("resume-game-button").style.display = "flex";
}

function playGame() {
  playing = true;
  interval = setInterval(updateGame, 500);
  document.getElementById("pause-game-button").style.display = "flex";
  document.getElementById("resume-game-button").style.display = "none";
}

function gameOver() {
  clearInterval(interval);
  interval = null;
  playing = false;
  document.getElementById("score-container").innerHTML = score.toString();
  document.getElementById("game-over-message").style.display = "flex";
  document.getElementById("pause-game-button").style.display = "none";
  document.getElementById("start-game-button").style.display = "flex";
}

function removeLines() {
  let linesCleared = 0;
  let multiplier = 0;
  for (let y = 0; y < Height; y++) {
    let complete = true;
    let colors = new Set();
    for (let x = 0; x < Width; x++) {
      complete = complete && blocks[x][y] != null;
      colors.add(blocks[x][y]);
    }
    if (complete) {
      linesCleared++;
      multiplier += colors.size;
      // Shift pieces down
      for (let y2 = y; y2 > 0; y2--) for (let x2 = 0; x2 < Width; x2++) blocks[x2][y2] = blocks[x2][y2 - 1];
      for (let x2 = 0; x2 < Width; x2++) blocks[x2][0] = null;
    }
  }
  score += Scores[linesCleared] * multiplier;
}

function pocket() {
  if (pocketPiece) {
    if (canPocket) {
      const tempPiece = currentPiece;
      tempPiece.y = 0;
      tempPiece.direction = Directions.Up;
      currentPiece = pocketPiece;
      pocketPiece = tempPiece;
      canPocket = false;
    }
  } else {
    currentPiece.y = 0;
    currentPiece.direction = Directions.Up;
    pocketPiece = currentPiece;
    cycleCurrentPiece();
  }
}

async function drawBoard() {
  try {
    await Excel.run(async context => {
      context.application.suspendScreenUpdatingUntilNextSync();

      const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
      // First, clear all previously drawn pieces
      activeWorksheet.getUsedRange().format.fill.clear();

      // Draw the game board
      let firstCell = activeWorksheet.getRangeByIndexes(1, 1, 1, 1);
      for (let x = 0; x < Width; x++) {
        for (let y = 0; y < Height; y++) {
          if (blocks[x][y] != null) firstCell.getOffsetRange(y, x).format.fill.color = BlockColorsMap[blocks[x][y]];
        }
      }
      eachBlock(currentPiece, (x: number, y: number) => {
        firstCell.getOffsetRange(y, x).format.fill.color = BlockColorsMap[currentPiece.pieceType.color];
      });

      // Draw the next piece
      firstCell = activeWorksheet.getRangeByIndexes(3, Width + 3, 1, 1);
      eachBlock(nextPiece, (x: number, y: number) => {
        firstCell.getOffsetRange(y, x - nextPiece.x).format.fill.color = BlockColorsMap[nextPiece.pieceType.color];
      });

      // Draw the pocket
      firstCell = activeWorksheet.getRangeByIndexes(8, Width + 3, 1, 1);
      if (pocketPiece) {
        eachBlock(pocketPiece, (x: number, y: number) => {
          firstCell.getOffsetRange(y, x - pocketPiece.x).format.fill.color =
            BlockColorsMap[pocketPiece.pieceType.color];
        });
      }

      // Score
      const scoreCell = activeWorksheet.getRangeByIndexes(1, Width + 2, 1, 1);
      scoreCell.values = [[`SCORE: ${score.toString()}`]];

      await context.sync();
    });
  } catch (e) {
    console.error(e);
  }
}

function setPieceInPlace() {
  // The piece can't move down anymore, so set it in place
  eachBlock(currentPiece, (x: string | number, y: string | number) => {
    blocks[x][y] = currentPiece.pieceType.color;
  });
  removeLines();
  cycleCurrentPiece();
  // Check if the game is over
  if (occupied(currentPiece)) gameOver();
  // Allow pocketing again
  canPocket = true;
}

async function updateGame() {
  drawBoard();
  if (!movePiece(currentPiece, Directions.Down)) {
    setPieceInPlace();
    await drawBoard();
  }
}

async function handleKeyPress(e: KeyboardEvent) {
  if (playing) {
    switch (e.code) {
      case "ArrowUp":
      case "KeyX":
        // Rotate 90 degrees clockwise
        rotatePiece(currentPiece, Directions.Left);
        break;
      case "ControlLeft":
      case "ControlRight":
      case "KeyZ":
        // Rotate 90 degrees conterclockwise
        rotatePiece(currentPiece, Directions.Right);
        break;
      case "KeyS":
        // Rotate 180 degrees
        rotatePiece(currentPiece, Directions.Right);
        rotatePiece(currentPiece, Directions.Right);
        break;
      case "ArrowLeft":
        // Move left
        movePiece(currentPiece, Directions.Left);
        break;
      case "ArrowRight":
        // Move right
        movePiece(currentPiece, Directions.Right);
        break;
      case "ArrowDown":
        // Soft drop
        movePiece(currentPiece, Directions.Down);
        break;
      case "Space":
        // Hard drop
        while (movePiece(currentPiece, Directions.Down));
        setPieceInPlace();
        break;
      case "ShiftLeft":
      case "ShiftRight":
      case "KeyC":
        // Pocket
        pocket();
        break;
      case "KeyP":
        // Pause
        pauseGame();
        break;
      case "KeyD":
        // Cycle colors
        currentPiece.pieceType.color = ((currentPiece.pieceType.color + 1) % 6) as BlockColors;
        break;
    }

    await drawBoard();
  } else if (e.code == "KeyP") {
    playGame();
  }
}

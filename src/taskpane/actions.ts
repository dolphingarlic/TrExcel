import { Directions } from "./constants";
import { Piece } from "./pieces";
import { occupied } from "./taskpane";

export function rotatePiece(currentPiece: Piece, rotateDirection: Directions) {
  const oldDirection = currentPiece.direction;
  const newDirection = (currentPiece.direction + rotateDirection) % 4;
  currentPiece.direction = newDirection;
  if (occupied(currentPiece) && !movePiece(currentPiece, Directions.Left) && !movePiece(currentPiece, Directions.Right))
    currentPiece.direction = oldDirection;
}

export function movePiece(currentPiece: Piece, moveDirection: Directions) {
  const prevX = currentPiece.x;
  const prevY = currentPiece.y;
  switch (moveDirection) {
    case Directions.Right:
      currentPiece.x++;
      break;
    case Directions.Left:
      currentPiece.x--;
      break;
    case Directions.Down:
      currentPiece.y++;
      break;
  }
  if (occupied(currentPiece)) {
    currentPiece.x = prevX;
    currentPiece.y = prevY;
    return false;
  }
  return true;
}

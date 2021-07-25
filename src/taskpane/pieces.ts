//-------------------------------------------------------------------------
// Game pieces
//
// blocks: each element represents a rotation of the piece (0, 90, 180, 270)
//         each element is a 16 bit integer where the 16 bits represent
//         a 4x4 set of blocks, e.g. JBlock.blocks[0] = 0x44C0
//
//             0100 = 0x4 << 3 = 0x4000
//             0100 = 0x4 << 2 = 0x0400
//             1100 = 0xC << 1 = 0x00C0
//             0000 = 0x0 << 0 = 0x0000
//                               ------
//                               0x44C0
//
//-------------------------------------------------------------------------

import { Directions, Width } from "./constants";

export enum BlockColors {
  Red,
  Orange,
  Yellow,
  Green,
  Blue,
  Purple
}

class PieceType {
  size: number;
  blocks: number[];
  color: BlockColors;

  constructor(size: number, blocks: number[]) {
    this.size = size;
    this.blocks = blocks;
    this.color = Math.floor(Math.random() * 6) as BlockColors;
  }
}

export class Piece {
  pieceType: PieceType;
  direction: Directions;
  x: number;
  y: number;

  constructor(pieceType: PieceType) {
    this.pieceType = pieceType;
    this.direction = Directions.Up;
    this.x = Math.floor(Math.random() * (Width - pieceType.size + 1));
    this.y = 0;
  }
}

const IBlock = new PieceType(4, [0x0f00, 0x2222, 0x00f0, 0x4444]);
const JBlock = new PieceType(3, [0x44c0, 0x8e00, 0x6440, 0x0e20]);
const LBlock = new PieceType(3, [0x4460, 0x0e80, 0xc440, 0x2e00]);
const OBlock = new PieceType(2, [0xcc00, 0xcc00, 0xcc00, 0xcc00]);
const SBlock = new PieceType(3, [0x06c0, 0x8c40, 0x6c00, 0x4620]);
const TBlock = new PieceType(3, [0x0e40, 0x4c40, 0x4e00, 0x4640]);
const ZBlock = new PieceType(3, [0x0c60, 0x4c80, 0xc600, 0x2640]);

//-----------------------------------------
// Start with 1 instance of each piece and
// pick randomly until the 'bag is empty'
//-----------------------------------------

let pieceTypes = [];

export function randomPiece() {
  if (pieceTypes.length == 0) pieceTypes = [IBlock, JBlock, LBlock, OBlock, SBlock, TBlock, ZBlock];

  const pieceType = pieceTypes.splice(Math.floor(Math.random() * pieceTypes.length), 1)[0];
  return new Piece(pieceType);
}

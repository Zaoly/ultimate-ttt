#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <windef.h>

// RESULT_TYPE: Win or lose (or even draw)
#define RES_UNKNOWN 0
#define RES_WIN 1
#define RES_DRAW 2
#define RES_LOSE 3

#define NULL_NODE 0 // Reserved

// PIECE_TYPE: X or O in Tic Tac Toe
#define NO_PIECE 0
#define PIECE_1 1
#define PIECE_2 2

#define RES_VAL_MIN -128
#define RES_VAL_MAX 127

typedef unsigned short POSITION_NUM; // Actually a 3*3 array, converted to number
typedef char RESULT_TYPE;
typedef unsigned long NODE_INDEX;
typedef char PIECE_TYPE;

typedef struct _NODE
{
	char move;                // Which square is moved in (0~8)
	POSITION_NUM position;
	char turnNo;              // Turn order, parity determining whose turn it is
	RESULT_TYPE result;
	char resultValue;         // Related to result and distance to the end of game
	NODE_INDEX firstSonIndex; // "Index" is one in nodeList (list of nodes)
	char sonDistance;
	NODE_INDEX parentIndex;
} NODE;

int nodeCount = 0;
NODE* nodeList = NULL; // Array to be allocated memory
char positionTable[19683] = {0};
char* maskData = "TicTacToeMoveTable ThisIsDevelopedByZaoly PleaseDoNotModifyMe SoAsNotToInterfereWithAutoMover ThankYouForCooperation Love Zaoly ";

int powint(int base, int exponent)
{
	int i = 0, result = 1;
	if (exponent < 0)
		return;
	for (; i < exponent; i ++)
		result *= base;
	return result;
}

/*

Position number & array conversion:

Array: (R = row, C = column; array[R][C])

 R\C  0 1 2

  0   A B C
  1   D E F
  2   G H I

Number:
 A * 1 + B * 3 + C * 9 + ... + I * 6561

*/

void positionNumToArray(POSITION_NUM positionNum, PIECE_TYPE positionArray[3][3])
{
	int row = 0, column = 0;
	for (; row < 3; row ++)
		for (column = 0; column < 3; column ++)
		{
			positionArray[row][column] = positionNum % 3;
			positionNum /= 3;
		}
}

POSITION_NUM positionArrayToNum(PIECE_TYPE positionArray[3][3])
{
	int row = 0, column = 0;
	POSITION_NUM result = 0;
	for (; row < 3; row ++)
		for (column = 0; column < 3; column ++)
			result += positionArray[row][column] * powint(3, row * 3 + column);
	return result;
}

POSITION_NUM moveToPosition(NODE_INDEX nodeIndex, char move)
{
	PIECE_TYPE positionArray[3][3] = {NO_PIECE};
	positionNumToArray(nodeList[nodeIndex].position, positionArray);
	if ((nodeList[nodeIndex].turnNo & 1) == 0)
		positionArray[move / 3][move % 3] = PIECE_1;
	else if ((nodeList[nodeIndex].turnNo & 1) == 1)
		positionArray[move / 3][move % 3] = PIECE_2;
	return positionArrayToNum(positionArray);
}

RESULT_TYPE judgePosition(NODE_INDEX nodeIndex)
{
	PIECE_TYPE posArr[3][3] = {NO_PIECE};
	int does1Win = 0, does2Win = 0;
	positionNumToArray(nodeList[nodeIndex].position, posArr);
	does1Win =
		posArr[0][0] == posArr[0][1] && posArr[0][1] == posArr[0][2] && posArr[0][2] == PIECE_1 ||
		posArr[1][0] == posArr[1][1] && posArr[1][1] == posArr[1][2] && posArr[1][2] == PIECE_1 ||
		posArr[2][0] == posArr[2][1] && posArr[2][1] == posArr[2][2] && posArr[2][2] == PIECE_1 ||
		posArr[0][0] == posArr[1][0] && posArr[1][0] == posArr[2][0] && posArr[2][0] == PIECE_1 ||
		posArr[0][1] == posArr[1][1] && posArr[1][1] == posArr[2][1] && posArr[2][1] == PIECE_1 ||
		posArr[0][2] == posArr[1][2] && posArr[1][2] == posArr[2][2] && posArr[2][2] == PIECE_1 ||
		posArr[0][0] == posArr[1][1] && posArr[1][1] == posArr[2][2] && posArr[2][2] == PIECE_1 ||
		posArr[0][2] == posArr[1][1] && posArr[1][1] == posArr[2][0] && posArr[2][0] == PIECE_1;
	/*
	
	Including:
	
		111    000    000    100    010    001    100    001
		000    111    000    100    010    001    010    010
		000    000    111    100    010    001    001    100
	
	*/
	
	does2Win =
		posArr[0][0] == posArr[0][1] && posArr[0][1] == posArr[0][2] && posArr[0][2] == PIECE_2 ||
		posArr[1][0] == posArr[1][1] && posArr[1][1] == posArr[1][2] && posArr[1][2] == PIECE_2 ||
		posArr[2][0] == posArr[2][1] && posArr[2][1] == posArr[2][2] && posArr[2][2] == PIECE_2 ||
		posArr[0][0] == posArr[1][0] && posArr[1][0] == posArr[2][0] && posArr[2][0] == PIECE_2 ||
		posArr[0][1] == posArr[1][1] && posArr[1][1] == posArr[2][1] && posArr[2][1] == PIECE_2 ||
		posArr[0][2] == posArr[1][2] && posArr[1][2] == posArr[2][2] && posArr[2][2] == PIECE_2 ||
		posArr[0][0] == posArr[1][1] && posArr[1][1] == posArr[2][2] && posArr[2][2] == PIECE_2 ||
		posArr[0][2] == posArr[1][1] && posArr[1][1] == posArr[2][0] && posArr[2][0] == PIECE_2;
	if (does1Win && !does2Win)      // First player wins
	{
		if ((nodeList[nodeIndex].turnNo & 1) == 0)      // Second player's turn
			return RES_LOSE;
		else if ((nodeList[nodeIndex].turnNo & 1) == 1) // First player's turn
			return RES_WIN;
	}
	else if (!does1Win && does2Win) // Second player wins
	{
		if ((nodeList[nodeIndex].turnNo & 1) == 0)      // Second player's turn
			return RES_WIN;
		else if ((nodeList[nodeIndex].turnNo & 1) == 1) // First player's turn
			return RES_LOSE;
	}
	else if
	(
		does1Win && does2Win ||     // Both players win
		posArr[0][0] != NO_PIECE && posArr[0][1] != NO_PIECE && posArr[0][2] != NO_PIECE &&
		posArr[1][0] != NO_PIECE && posArr[1][1] != NO_PIECE && posArr[1][2] != NO_PIECE &&
		posArr[2][0] != NO_PIECE && posArr[2][1] != NO_PIECE && posArr[2][2] != NO_PIECE // Board is full
	)
		return RES_DRAW;
	else // Game not over yet
		return RES_UNKNOWN;
}

NODE_INDEX addNode(char move, POSITION_NUM position, char turnNo, NODE_INDEX parentIndex) // Add item in nodeList
{
	NODE* newPtr = NULL;
	newPtr = realloc(nodeList, sizeof(NODE) * (nodeCount + 1)); // Allocate one more
	 // If there is no node, nodeList will equal NULL, but realloc(NULL, n) == malloc(n)
	if (newPtr == NULL)
		return NULL_NODE;
	nodeList = newPtr;
	nodeList[nodeCount].move = move;
	nodeList[nodeCount].position = position;
	nodeList[nodeCount].turnNo = turnNo;
	nodeList[nodeCount].result = RES_UNKNOWN;
	nodeList[nodeCount].resultValue = RES_VAL_MAX;
	nodeList[nodeCount].firstSonIndex = NULL_NODE;
	nodeList[nodeCount].sonDistance = 0;
	nodeList[nodeCount].parentIndex = parentIndex;
	nodeCount ++;
	return nodeCount - 1; // Return index of the added node
}

NODE_INDEX addSon(NODE_INDEX nodeIndex, char move)
{
	NODE_INDEX newNodeIndex = NULL_NODE;
	if (nodeIndex == NULL_NODE)
		return NULL_NODE;
	newNodeIndex = addNode(move, moveToPosition(nodeIndex, move), nodeList[nodeIndex].turnNo + 1, nodeIndex);
	if (newNodeIndex == NULL_NODE)
		return NULL_NODE;
	else if (nodeList[nodeIndex].firstSonIndex == NULL_NODE)
		nodeList[nodeIndex].firstSonIndex = newNodeIndex;
	else
		nodeList[nodeIndex].sonDistance ++;
	return newNodeIndex; // Return index of the added node
}

int addSons(NODE_INDEX nodeIndex)
{
	int row = 0, column = 0, successTimes = 0;
	PIECE_TYPE positionArray[3][3] = {NO_PIECE};
	if (nodeIndex == NULL_NODE)
		return 0;
	nodeList[nodeIndex].result = judgePosition(nodeIndex);
	if (nodeList[nodeIndex].result != RES_UNKNOWN) // Do not add when game is over
		return 0;
	positionNumToArray(nodeList[nodeIndex].position, positionArray);
	for (; row < 3; row ++)
		for (column = 0; column < 3; column ++)
			if (positionArray[row][column] == NO_PIECE) // Only move in empty squares
				if (addSon(nodeIndex, row * 3 + column) != NULL_NODE)
					successTimes ++;
				else
					printf("Error!\n");
	return successTimes;
}

char* resultNumToString(RESULT_TYPE result, char* string, size_t size)
{
	switch (result)
	{
		case RES_UNKNOWN: return strncpy(string, "UNKNOWN", size);
		case RES_WIN:     return strncpy(string, "WIN", size);
		case RES_DRAW:    return strncpy(string, "DRAW", size);
		case RES_LOSE:    return strncpy(string, "LOSE", size);
		default:          return NULL;
	}
}

void deleteNodes()
{
	free(nodeList);
	nodeCount = 0;
	nodeList = NULL;
}

int main()
{
	NODE_INDEX nodeIndex = NULL_NODE;
	char fileName[256] = "", resultString[8] = "";
	FILE* resultFile = NULL;
	int positionIndex = 0;
	addNode(0, 0, 0, NULL_NODE); // Null node
	addNode(0, 0, 0, NULL_NODE); // Main tree
	system("cmd /c title Tic Tac Toe Computer");
	printf("Press any key to start computing ...");
	system("pause>nul");
	printf("\nComputing ...\n");
	for (nodeIndex = 1; nodeIndex < nodeCount; nodeIndex ++)
	 // Do not worry, the loop will eventually complete though nodeCount increases as addSons() executes
		printf("%d", addSons(nodeIndex));
	printf("\nBacktracking ...\n");
	for (nodeIndex = nodeCount - 1; nodeIndex >= 1; nodeIndex --) // Backtrack game results
	{
		RESULT_TYPE newResult = RES_UNKNOWN;
		newResult = nodeList[nodeIndex].result;
		if (newResult == RES_WIN)       // My win means your defeat
			newResult = RES_LOSE;
		else if (newResult == RES_LOSE) // My defeat means your win
			newResult = RES_WIN;
		if (nodeList[nodeIndex].parentIndex != NULL_NODE)
			if (nodeList[nodeList[nodeIndex].parentIndex].result < newResult)
			 // Unknown < Win < Draw < Lose; take the maximum
				nodeList[nodeList[nodeIndex].parentIndex].result = newResult;
	}
	for (nodeIndex = nodeCount - 1; nodeIndex >= 1; nodeIndex --) // Compute distance to the end of game
	{
		char newResultValue = 0;
		if (nodeList[nodeIndex].firstSonIndex == NULL_NODE)
			if (nodeList[nodeIndex].result == RES_WIN)
				nodeList[nodeIndex].resultValue = RES_VAL_MIN;
			else if (nodeList[nodeIndex].result == RES_LOSE)
				nodeList[nodeIndex].resultValue = RES_VAL_MAX;
			else if (nodeList[nodeIndex].result == RES_DRAW)
				nodeList[nodeIndex].resultValue = 0;
		if (nodeList[nodeIndex].parentIndex != NULL_NODE)
		{
			newResultValue = nodeList[nodeIndex].resultValue;
			if (newResultValue > 0)
			{
				newResultValue --;
				newResultValue = RES_VAL_MAX - newResultValue + RES_VAL_MIN;
			}
			else if (newResultValue < 0)
			{
				newResultValue ++;
				newResultValue = RES_VAL_MAX - (newResultValue - RES_VAL_MIN);
			}
			if (nodeList[nodeList[nodeIndex].parentIndex].resultValue > newResultValue)
				nodeList[nodeList[nodeIndex].parentIndex].resultValue = newResultValue;
		}
	}
	for (nodeIndex = 1; nodeIndex < nodeCount; nodeIndex ++) // Turn move map into position map
	{
		if
		(
			positionTable[nodeList[nodeIndex].position] == 0 ||
			positionTable[nodeList[nodeIndex].position] == nodeList[nodeIndex].resultValue
		)
			positionTable[nodeList[nodeIndex].position] = nodeList[nodeIndex].resultValue;
		else
			printf("Error!\n");
	}
	printf("\nCompleted computing.");
	for (;;)
	{
		printf("\nOutput to file (text): ");
		gets(fileName);
		resultFile = fopen(fileName, "w");
		if (resultFile != NULL)
			break;
		printf("\nFail to open file.\n");
	}
	fprintf(resultFile, "id\tmove\tpos\ttrnNo\tres\tresStr\tresVal\tfrSonId\tsonDis\tpar\n");
	for (nodeIndex = 1; nodeIndex < nodeCount; nodeIndex ++)
	{
		resultNumToString(nodeList[nodeIndex].result, resultString, 8);
		fprintf(resultFile, "%lu\t%hd\t%hu\t%hd\t%hd\t%s\t%hd\t%lu\t%hd\t%lu\n",
			nodeIndex,
			(short)nodeList[nodeIndex].move,
			nodeList[nodeIndex].position,
			(short)nodeList[nodeIndex].turnNo,
			(short)nodeList[nodeIndex].result,
			resultString,
			(short)nodeList[nodeIndex].resultValue,
			nodeList[nodeIndex].firstSonIndex,
			(short)nodeList[nodeIndex].sonDistance,
			nodeList[nodeIndex].parentIndex
		);
	}
	fprintf(resultFile, "\npos\tresVal\n");
	for (; positionIndex < 19683; positionIndex ++)
		fprintf(resultFile, "%d\t%hd\n", positionIndex, (short)positionTable[positionIndex]);
	printf("\nOutput successfully.\n");
	fclose(resultFile);
	resultFile = NULL;
	for (;;)
	{
		printf("\nOutput to file (binary): ");
		gets(fileName);
		resultFile = fopen(fileName, "wb");
		if (resultFile != NULL)
			break;
		printf("\nFail to open file.\n");
	}
	for (positionIndex = 0; positionIndex < 19683; positionIndex ++)
		fputc(positionTable[positionIndex] ^ maskData[positionIndex & 127], resultFile);
	printf("\nOutput successfully.\n");
	fclose(resultFile);
	resultFile = NULL;
	deleteNodes();
	return 0;
}

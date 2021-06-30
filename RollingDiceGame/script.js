'use strict';

// Selecting elements
const player0EL = document.querySelector('.player--0');
const player1EL = document.querySelector('.player--1');
const score0El = document.querySelector('#score--0');
const score1EL = document.getElementById('score--1');
const current0EL = document.getElementById('current--0');
const current1EL = document.getElementById('current--1');

const diceEL = document.querySelector('.dice');
const btnNew = document.querySelector('.btn--new');
const btnRoll = document.querySelector('.btn--roll');
const btnHold = document.querySelector('.btn--hold');

const switchPlayer = function (){
    document.getElementById(`current--${activePlayer}`).textContent = 0;
    currentScore = 0;
    activePlayer = activePlayer === 0 ? 1 : 0;
    player0EL.classList.toggle('player--active');
    player1EL.classList.toggle('player--active');
};

let scores, currentScore, activePlayer, playing;


// Starting conditions
const init = function(){
     scores = [0,0];
     currentScore = 0;
     activePlayer = 0;
     playing = true;

    score0El.textContent = 0;
    score1EL.textContent = 0;
    current0EL.textContent = 0;
    current1EL.textContent = 0;

    diceEL.classList.add('hidden');
    player0EL.classList.remove('player-winner');
    player1EL.classList.remove('player--winner');
    player1EL.classList.remove('player--active');
    player0EL.classList.add('player--active');

};
init();

// Rolling dice functionality
btnRoll.addEventListener('click', function (){
    if(playing) {
        // 1. Generating a random dice roll
        const dice = Math.trunc(Math.random() * 6) + 1;

        // 2. Display dice
        diceEL.classList.remove('hidden');
        diceEL.src = `dice-${dice}.png`;

        // 3. Check for rolled 1
        if (dice !== 1) {
            // Add dice to current score
            currentScore += dice;
            // Dynamically building the id name
            document.getElementById(`current--${activePlayer}`).textContent = currentScore;

        } else {
            // switch to next player
            switchPlayer();
        }
    }

});

btnHold.addEventListener('click', function (){
    if(playing) {
        console.log('Hold Button');


        // 1. Add current score to active player's score
        scores[activePlayer] += currentScore;
        // scores[1] = scores[1] + currentScore;
        document.getElementById(`score--${activePlayer}`).textContent = scores[activePlayer];


        // 2. Check if player's scroe is >= 100
        if (scores[activePlayer] >= 20) {
            // Finish the game
            playing = false;
            document.querySelector(`.player--${activePlayer}`).classList.add('player--winner');
            document.querySelector(`.player--${activePlayer}`).classList.remove('player--active');
            diceEL.classList.add('hidden');

        } else {
            // Switch to the next player
            switchPlayer();
        }

    }

});
btnNew.addEventListener('click', init);




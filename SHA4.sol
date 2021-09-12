pragma solidity >=0.4.22 <0.6.0;

/*
 * Basic Shareholder's Agreement Smart contract
 */
 
contract ABC_ShareholderAgreement {
    address public sole_director;
    address public Mark_Joseph_Ferraretto = 0x5B38Da6a701c568545dCfcB03FcB875f56beddC4;
    mapping (address => uint) public share_registry;
    
    constructor() public {
        // Clause 1(a) of the Deed
        sole_director = Mark_Joseph_Ferraretto;
    }
    
    function issueShares(address receiver, uint amount) public {
        // Clause 2 of the Deed
        if (msg.sender != sole_director) return;
        share_registry[receiver] += amount;
    }
    
    function transferShares(address receiver, uint amount) public {
        // Clause 3 of the Deed
        if (share_registry[msg.sender] < amount) return;
        share_registry[msg.sender] -= amount;
        share_registry[receiver] += amount;
    }
}
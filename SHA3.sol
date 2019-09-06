pragma solidity >=0.4.22 <0.6.0;

/*
 * Basic Shareholder's Agreement Smart contract
 */
 
contract ABC_ShareholderAgreement {
    address public sole_director;
    mapping (address => uint) public share_registry;
    
    constructor() public {
        // Constructor.  This function is called only once when 
        // the contract is first created
        
        // We assume that the creator of the contract is
        // the sole director and we store his/her address
        sole_director = msg.sender;
    }
    
    function issueShares(address receiver, uint amount) public {
        // This part of the smart contract issues share 'certificates'
        // Only the sole director can issue shares/
        // The sole director has absolute discretion as to whom and throw
        // many shares he/she wishes to issue
        if (msg.sender != sole_director) return;
        share_registry[receiver] += amount;
    }
    
    function transferShares(address receiver, uint amount) public {
        // This part of the smart contract allows anyone who has shares
        // to transfer them to anyone else
        // - There are no restrictions on who the transferee can be
        // - The transferor must own the shares in order to transfer them
        if (share_registry[msg.sender] < amount) return;
        share_registry[msg.sender] -= amount;
        share_registry[receiver] += amount;
    }
}
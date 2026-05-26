CREATE PROCEDURE UpdateCompoundedContractCommissions as 

UPDATE    LevyContract

SET              LevyAmount = ProperCommission3

FROM         (SELECT     ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate, 

                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_, 

                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommissionRate, 

                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3

                       FROM          ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN

                                                  (SELECT     CAST(FLOOR(CAST(MAX(LotTDate) AS float)) AS datetime) AS LastDate

                                                    FROM          Lot

                                                    WHERE      (Deleted <> 1)) LastDateTransactions ON 

                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate = LastDateTransactions.LastDate) A

WHERE     LevyContract.LevyContract_DPA_ = A.LevyContract_DPA_



UPDATE    LevyContract

SET              LevyAmount = ProperAgentAmount

FROM         (SELECT     LevyContract.Contract_DPA_, LevyContract_1.LevyContract_DPA_ AS AgentLevyContract_DPA_, 

                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3 AS BrokerAmount, 

                                              LevyContract_1.LevyAmount AS AgentAmount, Commission.CommissionRate, 

                                              ROUND(Commission.CommissionRate / 100 * ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3, 2) 

                                              AS ProperAgentAmount

                       FROM          tbOrder INNER JOIN

                                              OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN

                                              Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN

                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN

                                              LevyContract ON 

                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_ = LevyContract.LevyContract_DPA_ INNER JOIN

                                              LevyContract LevyContract_1 ON LevyContract.Contract_DPA_ = LevyContract_1.Contract_DPA_ ON 

                                              Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN

                                              Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN

                                              Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ INNER JOIN

                                              Commission ON Agent.Commission_DPA_ = Commission.Commission_DPA_ INNER JOIN

                                                  (SELECT     CAST(FLOOR(CAST(MAX(LotTDate) AS float)) AS datetime) AS LastDate

                                                    FROM          Lot

                                                    WHERE      (Deleted <> 1)) LastDateTransactions ON CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) 

                                              = CAST(FLOOR(CAST(LastDateTransactions.LastDate AS float)) AS datetime)

                       WHERE      (LevyContract_1.SystemMaintained = 12) AND (Lot.Deleted <> 1)) A

WHERE     LevyContract.LevyContract_DPA_ = A.AgentLevyContract_DPA_

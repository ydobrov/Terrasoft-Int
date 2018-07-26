BEGIN
  MERGE INTO "SysAppQuestionnaireRight" Target
  USING (SELECT "Id"
         FROM "ApplicationQuestionnaire") Source
  ON (Target."RecordId" = Source."Id" AND
      Target."SysAdminUnitId" = '{7A53691B-926F-4B38-8C28-75D9DFC88F7C}' AND -- Клиентский менеджер
      Target."Operation" = 1 --> EntitySchemaRecRightOperation --> Изменение
      )
WHEN NOT MATCHED THEN INSERT ("RecordId", "SysAdminUnitId", "Operation", "RightLevel", "Position", "SourceId")
VALUES ( SOURCE."Id", '{7A53691B-926F-4B38-8C28-75D9DFC88F7C}', 1, 1, 0, '{8A248A03-E9A5-DF11-9989-485B39C18470}');
-- "RightLevel" = 1 -> SysEntitySchemaRecOprRightLvl -> Разрешено
-- "SourceId" = '{8A248A03-E9A5-DF11-9989-485B39C18470}' -> SysEntitySchemaRecRightSource -> Ручное управление
END;
ROLLBACK; -- COMMIT

BEGIN
  MERGE INTO "SysAppQuestionnaireRight" Target
  USING (SELECT "Id"
         FROM "ApplicationQuestionnaire") Source
  ON (Target."RecordId" = Source."Id" AND
      Target."SysAdminUnitId" = '{7A53691B-926F-4B38-8C28-75D9DFC88F7C}') -- Клиентский менеджер
  WHEN MATCHED THEN UPDATE SET Target."Operation" = 1, --> EntitySchemaRecRightOperation --> Изменение
    Target."RightLevel" = 1, --> SysEntitySchemaRecOprRightLvl --> Разрешено
    Target."Position" = 0,
    Target."SourceId" = '{8A248A03-E9A5-DF11-9989-485B39C18470}' --> SysEntitySchemaRecRightSource --> Ручное управление
  WHEN NOT MATCHED THEN INSERT ("RecordId", "SysAdminUnitId", "Operation", "RightLevel", "Position", "SourceId")
  VALUES (Source."Id", '{7A53691B-926F-4B38-8C28-75D9DFC88F7C}', 1, 1, 0, '{8A248A03-E9A5-DF11-9989-485B39C18470}');
END;
ROLLBACK; -- COMMIT